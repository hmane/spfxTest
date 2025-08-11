// net/http.ts
//
// HttpGateway: a tiny, context-bound wrapper around SPFx Http Clients
// -------------------------------------------------------------------
// Why use this instead of calling clients directly?
// - Centralized: timeout, retries, correlation-id, telemetry, and concurrency
// - Consistent: one return shape (HttpResult) for SP, AAD APIs, and Flows
// - Safer: optional idempotency for Flow triggers via replayId
//
// Quick usage (anywhere AFTER Context.setContext(...)):
//
//   const { http } = Context.getContext();
//
//   // 1) SharePoint REST (SPHttpClient)
//   const r1 = await http.sp("GET", `${ctx.pageContext.web.absoluteUrl}/_api/web`);
//   if (r1.ok) {
//     const web = JSON.parse(r1.body);
//   }
//
//   // 2) Azure AD–protected API (AadHttpClient)
//   //    - resource is your API's Application ID URI, e.g. "api://<client-id>"
//   //      or "https://contoso.onmicrosoft.com/contoso-api"
//   const r2 = await http.aad("api://11111111-2222-3333-4444-555555555555", "POST",
//                              "https://api.contoso.com/orders",
//                              { id: 123, qty: 2 });
//   if (!r2.ok) { /* handle error */ }
//
//   // 3) Power Automate Flow (convenience wrapper over http.sp)
//   //    - replayId prevents accidental double-submissions (idempotency)
//   const r3 = await http.flow({
//     url: "https://prod-XX.westus.logic.azure.com:443/workflows/.../invoke?api-version=2016-10-01",
//     body: { itemId: 42 },
//     replayId: "submit-nda-42"
//   });
//
// Notes:
// - Correlation Id: every request includes "X-Correlation-Id" (see Context.correlationId)
// - Timeout & Retries: configurable via Context options (here default 45s, 3 tries for 429/5xx)
// - Concurrency: a simple global semaphore limits in-flight requests (default 6). See globalSemaphore in throttle.ts
// - Telemetry: if App Insights is configured, dependency telemetry is sent automatically
//

import type { BaseComponentContext } from '@microsoft/sp-component-base';
import {
	AadHttpClientFactory,
	AadHttpClient,
	AadHttpClientConfiguration,
	HttpClient,
	IHttpClientOptions,
	HttpClientResponse,
	SPHttpClient,
	SPHttpClientResponse,
} from '@microsoft/sp-http';
import type { ApplicationInsights } from '@microsoft/applicationinsights-web';
import type { LoggerFacade } from '../logging/logging';
import { sleep } from '../utils/assert';
import type { HttpGatewayOptions, HttpResult } from '../utils/types';
import { globalSemaphore } from './throttle';

export class HttpGateway {
	private aadFactory: AadHttpClientFactory;
	private http: HttpClient; // exposed in case you need raw HttpClient later
	private spHttp: SPHttpClient;
	private maxConcurrent: number;

	constructor(private ctx: BaseComponentContext, private opts: HttpGatewayOptions) {
		this.aadFactory = this.ctx.serviceScope.consume(AadHttpClientFactory.serviceKey);
		this.http = this.ctx.httpClient;
		this.spHttp = (this.ctx as any).spHttpClient as SPHttpClient;
		this.maxConcurrent = this.opts.maxConcurrent ?? 6; // tune global concurrency here
	}

	/**
	 * Call SharePoint REST using SPHttpClient.
	 *
	 * @example
	 * const url = `${Context.getContext().webUrl}/_api/web/lists/getbytitle('Docs')/items?$top=50`;
	 * const res = await http.sp("GET", url);
	 * if (res.ok) {
	 *   const data = JSON.parse(res.body);
	 * }
	 */
	async sp(
		method: 'GET' | 'POST' | 'PATCH' | 'DELETE',
		url: string,
		body?: any,
		headers: Record<string, string> = {}
	): Promise<HttpResult> {
		return this._do('SharePoint', method, url, body, headers, async (reqInit) => {
			const res: SPHttpClientResponse = await this.spHttp.fetch(
				url,
				SPHttpClient.configurations.v1,
				reqInit
			);
			return { ok: res.ok, status: res.status, url: res.url, text: await res.text() };
		});
	}

	/**
	 * Call an Azure AD–protected API using AadHttpClient.
	 *
	 * @param resource Application ID URI for your API
	 *  - Examples: "api://<client-id>" OR "https://contoso.onmicrosoft.com/contoso-api"
	 * @example
	 * const res = await http.aad("api://1111-2222-3333-4444-5555", "POST",
	 *                            "https://api.contoso.com/orders", { id: 123 });
	 * const data = res.ok ? JSON.parse(res.body) : null;
	 */
	async aad(
		resource: string,
		method: 'GET' | 'POST' | 'PATCH' | 'DELETE',
		url: string,
		body?: any,
		headers: Record<string, string> = {}
	): Promise<HttpResult> {
		const client: AadHttpClient = await this.aadFactory.getClient(resource);
		return this._do('AAD', method, url, body, headers, async (reqInit) => {
			// Use AadHttpClient.configurations.v1 (note: NOT HttpClient.configurations.v1)
			const res: HttpClientResponse = await client.fetch(
				url,
				AadHttpClient.configurations.v1 as AadHttpClientConfiguration,
				reqInit
			);
			return { ok: res.ok, status: res.status, url: res.url, text: await res.text() };
		});
	}

	/**
	 * Trigger a Power Automate Flow exposed via HTTP endpoint.
	 *
	 * - Adds JSON Content-Type by default
	 * - Adds X-Correlation-Id for traceability
	 * - Optional replayId prevents double-submissions (Flow can dedupe by this header)
	 *
	 * @example
	 * const res = await http.flow({
	 *   url: "https://prod-XX.logic.azure.com/workflows/.../invoke?api-version=2016-10-01",
	 *   body: { itemId: 42 },
	 *   replayId: "submit-nda-42"
	 * });
	 */
	async flow(run: {
		url: string;
		method?: 'POST' | 'GET';
		body?: any;
		headers?: Record<string, string>;
		replayId?: string;
	}): Promise<HttpResult> {
		const method = run.method ?? 'POST';
		const headers = {
			'Content-Type': 'application/json',
			'X-Correlation-Id': this.opts.correlationId,
			...(run.replayId ? { 'X-Request-Id': run.replayId } : {}),
			...(run.headers || {}),
		};
		return this.sp(method, run.url, run.body, headers);
	}

	// -------------------------- internals --------------------------

	/**
	 * Core executor used by sp()/aad()/flow().
	 * Handles timeout, retries (429/5xx), telemetry, logging, and global concurrency.
	 *
	 * Return shape (HttpResult):
	 *   { ok: boolean, status: number, url: string, body: string, duration: ms }
	 *
	 * Tips:
	 * - Parse JSON: const data = res.ok ? JSON.parse(res.body) : null;
	 * - Handle errors: if (!res.ok) logger.error("API failed", { status: res.status, body: res.body });
	 */
	private async _do(
		kind: 'SharePoint' | 'AAD',
		method: string,
		url: string,
		body: any,
		headers: Record<string, string>,
		exec: (reqInit: IHttpClientOptions) => Promise<{
			ok: boolean;
			status: number;
			url: string;
			text: string;
		}>
	): Promise<HttpResult> {
		const start = performance.now();
		const ai = this.opts.ai as ApplicationInsights | undefined;
		const logger = this.opts.logger as LoggerFacade | undefined;

		// Global concurrency: acquire a "slot" before firing the request
		const release = await globalSemaphore.acquire();

		try {
			const reqInit: IHttpClientOptions = {
				method,
				headers: {
					Accept: 'application/json;odata=nometadata',
					'X-Correlation-Id': this.opts.correlationId,
					...headers,
				},
				// Strings are passed as-is; objects are JSON.stringify-ed
				body: body != null ? (typeof body === 'string' ? body : JSON.stringify(body)) : undefined,
			};

			let attempts = 0;
			const maxAttempts = Math.max(1, this.opts.retries ?? 1);
			let last: HttpResult = { ok: false, status: 0, url, body: '', duration: 0 };

			while (attempts < maxAttempts) {
				attempts++;
				try {
					// Enforce request timeout (default 45s if not provided)
					const {
						ok,
						status,
						url: finalUrl,
						text,
					} = await withTimeout(exec(reqInit), this.opts.timeoutMs ?? 45000);

					last = { ok, status, url: finalUrl, body: text, duration: performance.now() - start };

					// Application Insights dependency telemetry (if configured)
					// Note: SDK type names vary by version; we cast to any for portability.
					ai?.trackDependencyData({
						name: `${kind}:${method}`,
						target: url,
						data: url,
						duration: last.duration,
						success: ok,
						responseCode: status,
						dependencyTypeName: 'HTTP',
						properties: { attempts },
					} as any);

					// Verbose console/AppInsights log (if enabled in your logger)
					logger?.verbose('HTTP call', {
						kind,
						method,
						url,
						status,
						duration: last.duration,
						attempts,
					});

					// Simple retry policy: 429 or 5xx → backoff + retry
					if (!ok && (status === 429 || status >= 500) && attempts < maxAttempts) {
						await sleep(200 * attempts);
						continue;
					}
					return last;
				} catch (err: any) {
					// Network/timeout/unknown error
					last = {
						ok: false,
						status: 0,
						url,
						body: String(err?.message || err),
						duration: performance.now() - start,
					};

					ai?.trackException({
						exception: err instanceof Error ? err : new Error(String(err)),
						properties: { kind, method, url, attempts },
					});

					logger?.error('HTTP error', { kind, method, url, attempts, error: err?.message });

					if (attempts >= maxAttempts) return last;
					await sleep(200 * attempts);
				}
			}
			return last;
		} finally {
			// Release the concurrency slot
			release();
		}
	}
}

/**
 * Promise timeout helper.
 *
 * @example
 * const result = await withTimeout(fetch(...), 30000); // 30s
 */
async function withTimeout<T>(p: Promise<T>, ms: number): Promise<T> {
	let t: any;
	const timeout = new Promise<T>(
		(_, reject) => (t = setTimeout(() => reject(new Error(`Timeout ${ms}ms`)), ms))
	);
	try {
		return await Promise.race([p, timeout]);
	} finally {
		clearTimeout(t);
	}
}
