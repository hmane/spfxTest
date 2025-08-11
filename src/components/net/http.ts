// net/http.ts
//
// HttpGateway: a context-bound wrapper around SPFx Http Clients
// ------------------------------------------------------------
// Centralizes timeout, retries, correlation-id, Application Insights telemetry,
// and per-instance concurrency control. Provides escape hatches (raw clients)
// for advanced scenarios without breaking encapsulation.

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
import { Semaphore } from './throttle'; // per-instance semaphore

export class HttpGateway {
	private aadFactory: AadHttpClientFactory;
	private http: HttpClient;
	private spHttp: SPHttpClient;
	private sema: Semaphore;

	constructor(private ctx: BaseComponentContext, private opts: HttpGatewayOptions) {
		this.aadFactory = this.ctx.serviceScope.consume(AadHttpClientFactory.serviceKey);
		this.http = this.ctx.httpClient;
		this.spHttp = (this.ctx as any).spHttpClient as SPHttpClient;
		this.sema = new Semaphore(this.opts.maxConcurrent ?? 6); // use per-instance concurrency
	}

	/** Read-only access to the raw HttpClient for uncommon cases */
	public get rawHttp(): HttpClient {
		return this.http;
	}

	/** Read-only access to the raw SPHttpClient (SharePoint REST) */
	public get rawSpHttp(): SPHttpClient {
		return this.spHttp;
	}

	/** Helper to obtain an AadHttpClient for a resource when you need low-level control */
	public async getAadClient(resource: string): Promise<AadHttpClient> {
		return this.aadFactory.getClient(resource);
	}

	/**
	 * SharePoint REST via SPHttpClient
	 *
	 * @example
	 * const url = `${Context.getContext().webUrl}/_api/web`;
	 * const res = await http.sp("GET", url);
	 * const json = res.ok ? JSON.parse(res.body) : null;
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
	 * Azure AD–protected APIs via AadHttpClient
	 *
	 * @param resource Application ID URI for your API (e.g., "api://<client-id>")
	 * @example
	 * const r = await http.aad("api://1111-2222-3333-4444-5555", "POST", "https://api.contoso.com/orders", { id: 123 });
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
			const res: HttpClientResponse = await client.fetch(
				url,
				AadHttpClient.configurations.v1 as AadHttpClientConfiguration,
				reqInit
			);
			return { ok: res.ok, status: res.status, url: res.url, text: await res.text() };
		});
	}

	/**
	 * Power Automate Flow convenience
	 * - Adds JSON content type
	 * - Adds X-Correlation-Id
	 * - Optional replayId header for idempotency
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
	 * Core executor used by sp()/aad()/flow()
	 * - per-instance concurrency via Semaphore
	 * - timeout + simple retries (429/5xx)
	 * - AI dependency telemetry (if configured)
	 * - structured return shape (HttpResult)
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

		// acquire a concurrency slot (per-instance)
		const release = await this.sema.acquire();

		try {
			const reqInit: IHttpClientOptions = {
				method,
				headers: {
					Accept: 'application/json;odata=nometadata',
					'X-Correlation-Id': this.opts.correlationId,
					...headers,
				},
				body: body != null ? (typeof body === 'string' ? body : JSON.stringify(body)) : undefined,
			};

			let attempts = 0;
			const maxAttempts = Math.max(1, this.opts.retries ?? 1);
			let last: HttpResult = { ok: false, status: 0, url, body: '', duration: 0 };

			while (attempts < maxAttempts) {
				attempts++;
				try {
					// enforce timeout per call
					const {
						ok,
						status,
						url: finalUrl,
						text,
					} = await withTimeout(exec(reqInit), this.opts.timeoutMs ?? 45_000);

					last = { ok, status, url: finalUrl, body: text, duration: performance.now() - start };

					// App Insights dependency telemetry (SDK types vary — cast to any for portability)
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

					logger?.verbose('HTTP call', {
						kind,
						method,
						url,
						status,
						duration: last.duration,
						attempts,
					});

					// backoff retry on 429/5xx
					if (!ok && (status === 429 || status >= 500) && attempts < maxAttempts) {
						await sleep(200 * attempts);
						continue;
					}
					return last;
				} catch (err: any) {
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
			release(); // free concurrency slot
		}
	}
}

/** Promise timeout helper */
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
