// utils/types.ts
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

export interface ContextOptions {
	logLevel?: LogLevel;
	defaultCache?: CacheTTL;
	timeoutMs?: number;
	appInsightsKey?: string;
	aiRoleName?: string;
	enableTelemetry?: boolean;
	enableDiagnosticsSink?: boolean;
	componentName?: string;
}

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

	http: HttpGateway;
	links: LinksBound;
	logger: LoggerFacade;

	forWeb(webUrl: string, cache?: CacheTTL): SPFI;
	with<T>(cache: CacheTTL, fn: (sp: SPFI, graph: GraphFI) => Promise<T>): Promise<T>;
}

export interface HttpGatewayOptions {
	timeoutMs?: number;
	retries?: number;
	correlationId: string;
	logger?: LoggerFacade;
	ai?: ApplicationInsights;
	maxConcurrent?: number;
}

export interface HttpResult {
	ok: boolean;
	status: number;
	url: string;
	body: string;
	duration: number;
}
