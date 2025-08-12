/**
 * context/pnpjs-config.ts
 * --------------------------------------------------------------
 * Single source of truth for SPFx context + PnPjs (SP & Graph).
 *
 * Call once in your host's onInit():
 *   Context.setContext(this.context, { componentName: "My WebPart", ...options })
 *
 * Then anywhere:
 *   const { sp, graph, http, links, logger, env, isProdSite, buildMode, envBadge } = Context.getContext();
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
import type { CacheTTL, ContextOptions, EnvName, SPContext } from '../utils/types';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';

// ---------- constants ----------
const TTL_SHORT_MS = 5 * 60 * 1000; // 5 minutes
const TTL_LONG_MS = 24 * 60 * 60 * 1000; // 24 hours
const REQ_TIMEOUT_MS = 45_000; // 45 seconds default

// ---------- helpers ----------
function newCorrelationId(): string {
	return `spctx-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

/** v3 expects Date from expireFunc */
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

function getBuildMode(): 'production' | 'development' | 'test' | 'unknown' {
	try {
		// 1) SPFx explicit environment
		if (Environment.type === EnvironmentType.Local) return 'development'; // gulp serve
		if (Environment.type === EnvironmentType.Test) return 'test'; // unit/integration harnesses

		// 2) Browser hints (hosted but dev-like)
		const loc = typeof window !== 'undefined' ? window.location : undefined;
		const path = (loc?.pathname || '').toLowerCase();
		const host = (loc?.hostname || '').toLowerCase();

		// Localhost or workbench typically indicates dev
		if (host === 'localhost' || path.includes('/_layouts/15/workbench.aspx')) return 'development';

		// 3) Default: production (bundled, deployed)
		return 'production';
	} catch {
		return 'unknown';
	}
}

/**
 * Detect site env by your path rules:
 *  - /debug{SITE_NAME} → dev
 *  - /dev{SITE_NAME}   → dev
 *  - /uat{SITE_NAME}   → uat
 *  - /{SITE_NAME}      → prod
 */
function detectSiteEnv(webAbsoluteUrl: string, webServerRelativeUrl: string): EnvName {
	try {
		const path = (webServerRelativeUrl || new URL(webAbsoluteUrl).pathname || '/').toLowerCase();
		const seg = path.startsWith('/') ? path.slice(1) : path;
		const first = seg.split('/')[0] || '';
		if (first.startsWith('debug')) return 'dev';
		if (first.startsWith('dev')) return 'dev';
		if (first.startsWith('uat')) return 'uat';
		return 'prod';
	} catch {
		return 'prod';
	}
}

/** Optional query override: ?spctxenv=dev|uat|prod */
function readEnvOverride(): EnvName | undefined {
	try {
		const v = new URLSearchParams(window.location.search).get('spctxenv');
		if (!v) return undefined;
		const x = v.toLowerCase();
		return x === 'dev' || x === 'uat' || x === 'prod' ? (x as EnvName) : undefined;
	} catch {
		return undefined;
	}
}

/** Small UI helper for an environment badge */
function makeEnvBadge(
	env: EnvName,
	isProdBuild: boolean
): () => { text: string; color: string; tooltip: string } {
	return () => {
		switch (env) {
			case 'dev':
				return {
					text: 'DEV',
					color: '#6D28D9',
					tooltip: `${isProdBuild ? 'Prod build' : 'Dev build'} on DEV site`,
				};
			case 'uat':
				return {
					text: 'UAT',
					color: '#F59E0B',
					tooltip: `${isProdBuild ? 'Prod build' : 'Dev build'} on UAT site`,
				};
			case 'prod':
			default:
				return {
					text: 'PROD',
					color: isProdBuild ? '#059669' : '#9CA3AF',
					tooltip: `${isProdBuild ? 'Prod build' : 'Dev build'} on PROD site`,
				};
		}
	};
}

// ---------- singleton hub ----------
class ContextHub {
	private static _instance: ContextHub | null = null;
	static get instance() {
		return this._instance ?? (this._instance = new ContextHub());
	}

	private _ctx?: SPContext;

	isReady() {
		return !!this._ctx;
	}
	getContext(): SPContext {
		assert(!!this._ctx, 'ContextHub not initialized. Call setContext() in onInit() first.');
		return this._ctx!;
	}

	setContext(ctx: WebPartContext | BaseComponentContext, options?: ContextOptions): SPContext {
		const pc = (ctx as any).pageContext as PageContext;
		assert(!!pc, 'ContextHub.setContext: missing pageContext');

		const webAbs = pc?.web?.absoluteUrl ?? '';
		const webRel = pc?.web?.serverRelativeUrl ?? '';

		// ----- env & build detection -----
		const envOverride = readEnvOverride();
		const siteEnv: EnvName = envOverride ?? detectSiteEnv(webAbs, webRel);
		const isProdSite = siteEnv === 'prod';
		const buildMode = getBuildMode();
		const isProdBuild = buildMode === 'production';

		// ----- dev URL overrides + options -----
		const ov = readUrlOverrides();
		const logLevel =
			ov.logLevel ?? options?.logLevel ?? (isProdSite ? LogLevel.Warning : LogLevel.Info);
		const defaultCache: CacheTTL = ov.defaultCache ?? options?.defaultCache ?? 'short';
		const timeoutMs = ov.timeoutMs ?? options?.timeoutMs ?? REQ_TIMEOUT_MS;
		const enableTelemetry =
			ov.enableTelemetry ?? options?.enableTelemetry ?? !!options?.appInsightsKey;

		Logger.activeLogLevel = logLevel;

		// ----- logger (+ sinks) -----
		const correlationId = newCorrelationId();
		const logger = createLogger({
			level: logLevel,
			defaultCategory: options?.componentName ?? 'App',
			enrich: () => ({
				correlationId,
				env: siteEnv,
				isProdSite,
				buildMode,
				isProdBuild,
				siteUrl: pc?.site?.absoluteUrl || '',
				webUrl: webAbs || '',
				user: pc?.user?.loginName || '',
			}),
			redactor: redactDeep,
		});
		logger.addSink(new ConsoleSink());
		if (options?.enableDiagnosticsSink) logger.addSink(new DiagnosticsSink());

		let ai = undefined as ReturnType<typeof initAppInsights> | undefined;
		if (enableTelemetry && options?.appInsightsKey) {
			ai = initAppInsights(options.appInsightsKey, options.aiRoleName ?? 'SPFxApp');
			logger.addSink(new AppInsightsSink(ai));
			bridgePnPLoggerToSinks(logger);
		}

		// ----- PnP instances + cache flavors -----
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

		// ----- HTTP gateway -----
		const http = new HttpGateway(ctx, {
			timeoutMs,
			retries: 3,
			correlationId,
			logger,
			ai,
			maxConcurrent: 6,
		});

		// ----- finalize context -----
		const sc: SPContext = {
			context: ctx,
			pageContext: pc,

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

			logLevel,
			correlationId,

			siteUrl: pc?.site?.absoluteUrl ?? '',
			webUrl: pc?.web?.absoluteUrl ?? '',
			webRelativeUrl: pc?.web?.serverRelativeUrl ?? '',
			webAbsoluteUrl: webAbs,
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

			http,
			links: buildLinks({ webAbsoluteUrl: webAbs, webRelativeUrl: webRel }),
			logger,

			env: siteEnv,
			isProdSite,
			buildMode,
			isProdBuild,
			envBadge: makeEnvBadge(siteEnv, isProdBuild),

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

		this._ctx = sc;

		if (process.env.NODE_ENV !== 'production') {
			(window as any).__spctx = {
				get: () => this._ctx,
				env: siteEnv,
				buildMode,
				logLevel,
				defaultCache,
				timeoutMs,
			};
		}

		logger.banner(`Context ready • ${pc?.web?.title || 'Web'}`, {
			site: sc.siteUrl,
			web: sc.webUrl,
			user: sc.currentUserLoginName,
			env: siteEnv,
			buildMode,
			correlationId,
		});

		return sc;
	}
}

export const Context = ContextHub.instance;
export const getSp = () => Context.getContext().sp;
export const getGraph = () => Context.getContext().graph;
