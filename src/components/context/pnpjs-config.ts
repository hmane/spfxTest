/**
 * context/pnpjs-config.ts
 * --------------------------------------------------------------
 * Single source of truth for SPFx context + PnPjs (SP & Graph).
 * - Call `Context.setContext(this.context, options?)` ONCE in onInit()
 * - Anywhere else: `const ctx = Context.getContext();`
 *   - ctx.sp / ctx.graph  → default PnP instances
 *   - cache flavors       → spNoCache/spShortCached/spLongCached/spPessimisticRefresh (and Graph equivalents)
 *   - ctx.http            → SP/AAD/Flow HTTP gateway (timeout/retries/AI telemetry/concurrency)
 *   - ctx.links           → context-bound link builders (no webUrl param needed)
 *   - ctx.logger          → shared logger (Console + optional Diagnostics + optional AppInsights)
 *
 * URL overrides for dev (optional):
 *   ?spctxlog=0..4|verbose|info|warning|error
 *   ?spctxcache=none|short|long|pess
 *   ?spctxtimeout=45000
 *   ?spctxtelemetry=1|0
 */

import { LogLevel, Logger } from '@pnp/logging';
import { spfi, SPFI, SPFx as SPFxBind } from '@pnp/sp';
import { graphfi, GraphFI, SPFx as GraphSPFxBind } from '@pnp/graph';
import { Caching, CachingPessimisticRefresh } from '@pnp/queryable';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { BaseComponentContext } from '@microsoft/sp-component-base';
import type { PageContext } from '@microsoft/sp-page-context';

import {
	AppInsightsSink,
	bridgePnPLoggerToSinks,
	createLogger,
	ConsoleSink,
	DiagnosticsSink,
} from '../logging/logging';
import { redactDeep } from '../logging/redactor';
import { initAppInsights } from '../net/telemetry';
import { HttpGateway } from '../net/http';
import { buildLinks } from '../utils/links';
import { assert } from '../utils/assert';
import type { CacheTTL, ContextOptions, SPContext } from '../utils/types';

// ---------- constants (cache TTLs, default timeout) ----------
const TTL_SHORT_MS = 5 * 60 * 1000; // 5 minutes
const TTL_LONG_MS = 24 * 60 * 60 * 1000; // 1 day
const REQ_TIMEOUT_MS = 45_000; // 45 seconds default

// ---------- helpers ----------
/** Correlation id to stitch logs/requests across modules */
function newCorrelationId(): string {
	return `spctx-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

/** v3 expects a Date, not a number, for expireFunc */
function cacheBehavior(ttlMs: number) {
	return Caching({
		store: 'local',
		keyFactory: (u) => u.toLowerCase(),
		expireFunc: () => new Date(Date.now() + ttlMs),
	});
}
function pessimisticBehavior(ttlMs: number) {
	return CachingPessimisticRefresh({
		store: 'local',
		keyFactory: (u) => u.toLowerCase(),
		expireFunc: () => new Date(Date.now() + ttlMs),
	});
}

/** Support both numeric and string log levels via querystring */
function parseLogLevel(raw?: string | null): LogLevel | undefined {
	if (!raw) return undefined;
	const n = Number(raw);
	if (Number.isFinite(n)) return Math.max(0, Math.min(4, n)) as LogLevel;
	switch (raw.toLowerCase()) {
		case 'verbose':
			return LogLevel.Verbose;
		case 'info':
			return LogLevel.Info;
		case 'warning':
			return LogLevel.Warning;
		case 'error':
			return LogLevel.Error;
		default:
			return undefined;
	}
}

/** Read dev URL overrides (safe to ignore if absent) */
function readUrlOverrides(): Partial<
	Pick<ContextOptions, 'logLevel' | 'defaultCache' | 'timeoutMs' | 'enableTelemetry'>
> {
	try {
		const usp = new URLSearchParams(window.location.search);
		const logLevel = parseLogLevel(usp.get('spctxlog'));
		const cacheRaw = usp.get('spctxcache'); // none|short|long|pess
		const timeoutMs = usp.get('spctxtimeout')
			? Math.max(5000, Number(usp.get('spctxtimeout')))
			: undefined;
		const enableTelemetry =
			usp.get('spctxtelemetry') != null ? usp.get('spctxtelemetry') !== '0' : undefined;

		const defaultCache: CacheTTL | undefined =
			cacheRaw === 'none'
				? 'none'
				: cacheRaw === 'long'
				? 'long'
				: cacheRaw === 'pess'
				? 'pessimistic'
				: cacheRaw === 'short'
				? 'short'
				: undefined;

		return { logLevel, defaultCache, timeoutMs, enableTelemetry };
	} catch {
		return {};
	}
}

// ---------- singleton hub ----------
class ContextHub {
	private static _instance: ContextHub | null = null;
	static get instance() {
		return this._instance ?? (this._instance = new ContextHub());
	}

	private _ctx?: SPContext;

	/** Has setContext been called? */
	isReady() {
		return !!this._ctx;
	}

	/** Return the initialized context (throws if not ready) */
	getContext(): SPContext {
		assert(!!this._ctx, 'ContextHub not initialized. Call setContext() in onInit() first.');
		return this._ctx!;
	}

	/**
	 * Initialize once in your host's onInit():
	 *   Context.setContext(this.context, { appInsightsKey: "...", enableDiagnosticsSink: true })
	 */
	setContext(ctx: WebPartContext | BaseComponentContext, options?: ContextOptions): SPContext {
		const pc = (ctx as any).pageContext as PageContext;
		assert(!!pc, 'ContextHub.setContext: missing pageContext');

		// 1) resolve options + url overrides
		const ov = readUrlOverrides();
		const logLevel = ov.logLevel ?? options?.logLevel ?? LogLevel.Warning;
		const defaultCache: CacheTTL = ov.defaultCache ?? options?.defaultCache ?? 'short';
		const timeoutMs = ov.timeoutMs ?? options?.timeoutMs ?? REQ_TIMEOUT_MS;
		const enableTelemetry =
			ov.enableTelemetry ?? options?.enableTelemetry ?? !!options?.appInsightsKey;

		// apply global PnP log level (affects @pnp/logging subscribers)
		Logger.activeLogLevel = logLevel;

		// 2) bootstrap logger (+ sinks)
		const correlationId = newCorrelationId();
		const logger = createLogger({
			level: logLevel,
			defaultCategory: options?.componentName ?? 'App',
			enrich: () => ({
				correlationId,
				siteUrl: pc?.site?.absoluteUrl || '',
				webUrl: pc?.web?.absoluteUrl || '',
				user: pc?.user?.loginName || '',
			}),
			redactor: redactDeep,
		});
		// Always log to browser console
		logger.addSink(new ConsoleSink());
		// Optional: SPFx Developer Console (enable via option)
		if (options?.enableDiagnosticsSink) logger.addSink(new DiagnosticsSink());

		// Optional: Application Insights sink
		let ai = undefined as ReturnType<typeof initAppInsights> | undefined;
		if (enableTelemetry && options?.appInsightsKey) {
			ai = initAppInsights(options.appInsightsKey, options.aiRoleName ?? 'SPFxApp');
			logger.addSink(new AppInsightsSink(ai));
			// Bridge legacy PnP logger to our sinks (so old code logs consistently)
			bridgePnPLoggerToSinks(logger);
		}

		// 3) create PnP instances (SP & Graph) + caching “flavors”
		//    (no custom queryable behaviors so it compiles cleanly with @pnp/sp 3.20.x)
		const spBase = spfi().using(SPFxBind(ctx));
		const spNoCache = spBase;
		const spShortCached = spBase.using(cacheBehavior(TTL_SHORT_MS));
		const spLongCached = spBase.using(cacheBehavior(TTL_LONG_MS));
		const spPessimisticRefresh = spBase.using(pessimisticBehavior(TTL_LONG_MS));

		const spDefault =
			defaultCache === 'none'
				? spNoCache
				: defaultCache === 'long'
				? spLongCached
				: defaultCache === 'pessimistic'
				? spPessimisticRefresh
				: spShortCached;

		const gBase = graphfi().using(GraphSPFxBind(ctx));
		const graphNoCache = gBase;
		const graphShortCached = gBase.using(cacheBehavior(TTL_SHORT_MS));
		const graphLongCached = gBase.using(cacheBehavior(TTL_LONG_MS));
		const graphPessimisticRefresh = gBase.using(pessimisticBehavior(TTL_LONG_MS));

		const graphDefault =
			defaultCache === 'none'
				? graphNoCache
				: defaultCache === 'long'
				? graphLongCached
				: defaultCache === 'pessimistic'
				? graphPessimisticRefresh
				: graphShortCached;

		// 4) HTTP gateway (central timeout/retry/concurrency + AI dependency telemetry)
		const http = new HttpGateway(ctx, {
			timeoutMs,
			retries: 3,
			correlationId,
			logger,
			ai,
			maxConcurrent: 6,
		});

		// 5) Build the fully-bound context object
		const sc: SPContext = {
			context: ctx,
			pageContext: pc,

			// PnP instances
			sp: spDefault,
			graph: graphDefault,

			spNoCache,
			spShortCached,
			spLongCached,
			spPessimisticRefresh,

			graphNoCache,
			graphShortCached,
			graphLongCached,
			graphPessimisticRefresh,

			// env/meta
			logLevel,
			correlationId,

			siteUrl: pc?.site?.absoluteUrl ?? '',
			webUrl: pc?.web?.absoluteUrl ?? '',
			webRelativeUrl: pc?.web?.serverRelativeUrl ?? '',
			webAbsoluteUrl: pc?.web?.absoluteUrl ?? '',
			siteId: (pc as any)?.site?.id?.toString?.() ?? (pc as any)?.site?.id ?? '',
			webId: (pc as any)?.web?.id?.toString?.() ?? (pc as any)?.web?.id ?? '',
			webTitle: pc?.web?.title ?? '',
			cultureName: (pc as any)?.cultureInfo?.currentCultureName,
			uiLang: (pc as any)?.cultureInfo?.currentUICultureName,
			aadTenantId: (pc as any)?.azureActiveDirectoryInfo?.tenantId ?? '',
			currentUserLoginName: pc?.user?.loginName,
			currentUserDisplayName: (pc as any)?.user?.displayName,
			isTeams: !!(pc as any)?.legacyPageContext?.isTeamsContext,
			isClassicPage: !!(pc as any)?.legacyPageContext?.isClassicPage,

			// helpers
			http,
			links: buildLinks({
				webAbsoluteUrl: pc?.web?.absoluteUrl ?? '',
				webRelativeUrl: pc?.web?.serverRelativeUrl ?? '',
			}),
			logger,

			/**
			 * Create a new SPFI scoped to another Web URL, with a chosen cache flavor.
			 * Useful for cross-site reads without re-initializing the hub.
			 */
			forWeb: (webUrl: string, cache: CacheTTL = 'short'): SPFI => {
				const base = spfi(webUrl).using(SPFxBind(ctx));
				switch (cache) {
					case 'none':
						return base;
					case 'long':
						return base.using(cacheBehavior(TTL_LONG_MS));
					case 'pessimistic':
						return base.using(pessimisticBehavior(TTL_LONG_MS));
					default:
						return base.using(cacheBehavior(TTL_SHORT_MS));
				}
			},

			/**
			 * Temporarily switch both SP & Graph to a specific cache flavor for a block of work.
			 * Example:
			 *   const result = await Context.getContext().with("long", async (sp, graph) => {
			 *     const items = await sp.web.lists.getByTitle("Catalog").items.top(500)();
			 *     const me = await graph.me();
			 *     return { items, me };
			 *   });
			 */
			with: async <T>(cache: CacheTTL, fn: (sp: SPFI, graph: GraphFI) => Promise<T>) => {
				const spSel =
					cache === 'none'
						? spNoCache
						: cache === 'long'
						? spLongCached
						: cache === 'pessimistic'
						? spPessimisticRefresh
						: spShortCached;

				const graphSel =
					cache === 'none'
						? graphNoCache
						: cache === 'long'
						? graphLongCached
						: cache === 'pessimistic'
						? graphPessimisticRefresh
						: graphShortCached;

				return fn(spSel, graphSel);
			},
		};

		// 6) stash + dev handle
		this._ctx = sc;
		if (process.env.NODE_ENV !== 'production') {
			(window as any).__spctx = { get: () => this._ctx, logLevel, defaultCache, timeoutMs };
		}

		// 7) hello world log
		logger.banner(`Context ready • ${pc?.web?.title || 'Web'}`, {
			site: sc.siteUrl,
			web: sc.webUrl,
			user: sc.currentUserLoginName,
			correlationId,
		});

		return sc;
	}
}

// ---------- public exports ----------
export const Context = ContextHub.instance;
export const getSp = () => Context.getContext().sp;
export const getGraph = () => Context.getContext().graph;
