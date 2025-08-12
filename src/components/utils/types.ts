// utils/types.ts
// Shared types for the SPFx Context Hub

import type { LogLevel } from '@pnp/logging';
import type { ApplicationInsights } from '@microsoft/applicationinsights-web';
import type { SPFI } from '@pnp/sp';
import type { GraphFI } from '@pnp/graph';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { BaseComponentContext } from '@microsoft/sp-component-base';
import type { PageContext } from '@microsoft/sp-page-context';
import type { HttpGateway } from '../net/http';
import type { LoggerFacade } from '../logging/logging';
import type { LinksBound } from './links';

export type CacheTTL = 'none' | 'short' | 'long' | 'pessimistic';
export type EnvName = 'dev' | 'uat' | 'prod';

export interface ContextOptions {
	/** Global log verbosity (default: Warning on prod sites, Info elsewhere) */
	logLevel?: LogLevel;
	/** Default cache flavor when none is specified (default: "short") */
	defaultCache?: CacheTTL;
	/** HTTP timeout for http.sp/aad/flow (default: 45000 ms) */
	timeoutMs?: number;
	/** App Insights key; telemetry auto-enables if present unless enableTelemetry=false */
	appInsightsKey?: string;
	/** Role name to tag in App Insights (default: "SPFxApp") */
	aiRoleName?: string;
	/** Force telemetry on/off. Default: on iff appInsightsKey present; otherwise off */
	enableTelemetry?: boolean;
	/** Also log to SPFx Developer Console */
	enableDiagnosticsSink?: boolean;
	/** Friendly component name that appears in logs/telemetry */
	componentName?: string;
}

export interface HttpGatewayOptions {
	timeoutMs?: number;
	retries?: number;
	correlationId: string;
	logger?: LoggerFacade;
	ai?: ApplicationInsights;
	/** Per-instance concurrency limit (default: 6) */
	maxConcurrent?: number;
}

export interface HttpResult {
	ok: boolean;
	status: number;
	url: string;
	body: string;
	duration: number;
}

export interface EnvBadge {
	/** Badge text to show in UI ("DEV", "UAT", "PROD") */
	text: string;
	/** Hex color (background) suggestion for the badge */
	color: string;
	/** Title/tooltip explaining build/site */
	tooltip: string;
}

/** The full, ready-to-use context your app consumes everywhere */
export interface SPContext {
	context: WebPartContext | BaseComponentContext;
	pageContext: PageContext;

	sp: SPFI;
	graph: GraphFI;

	spNoCache: SPFI;
	spShortCached: SPFI;
	spLongCached: SPFI;
	spPessimisticRefresh: SPFI;

	graphNoCache: GraphFI;
	graphShortCached: GraphFI;
	graphLongCached: GraphFI;
	graphPessimisticRefresh: GraphFI;

	logLevel: LogLevel;
	correlationId: string;

	siteUrl: string;
	webUrl: string;
	webRelativeUrl: string;
	webAbsoluteUrl: string;
	siteId: string;
	webId: string;
	webTitle: string;
	cultureName?: string;
	uiLang?: string;
	aadTenantId?: string;
	currentUserLoginName?: string;
	currentUserDisplayName?: string;
	isTeams: boolean;
	isClassicPage: boolean;

	/** Centralized HTTP gateway (SP, AAD, Flow) */
	http: HttpGateway;
	/** Context-bound link builders */
	links: LinksBound;
	/** Shared, enriched logger (console + optional Diagnostics + optional AppInsights) */
	logger: LoggerFacade;

	/** Site environment resolved from the URL (dev/uat/prod) */
	env: EnvName;
	/** True if env === "prod" */
	isProdSite: boolean;
	/** Webpack build mode from process.env.NODE_ENV */
	buildMode: 'production' | 'development' | 'test' | 'unknown';
	/** True if buildMode === "production" */
	isProdBuild: boolean;
	/** Small helper for showing an environment badge in your UI */
	envBadge(): EnvBadge;

	/** Create an SPFI for another web URL with a chosen cache flavor */
	forWeb(webUrl: string, cache?: CacheTTL): SPFI;

	/** Temporarily switch both SP & Graph to a flavor for a block of work */
	with<T>(cache: CacheTTL, fn: (sp: SPFI, graph: GraphFI) => Promise<T>): Promise<T>;
}
