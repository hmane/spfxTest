import { WebPartContext } from '@microsoft/sp-webpart-base';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { ExtensionContext } from '@microsoft/sp-extension-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { spfi, SPFI, SPFx } from '@pnp/sp';
import { graphfi, GraphFI } from '@pnp/graph';
import { LogLevel, PnPLogging } from '@pnp/logging';
import { Caching, ICachingProps } from '@pnp/queryable';

// Import required PnP modules
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import '@pnp/sp/batching';
import '@pnp/sp/site-users/web';
import '@pnp/sp/profiles';
import '@pnp/sp/search';
import '@pnp/graph/users';
import '@pnp/graph/groups';

export type SPFxContext = WebPartContext | ExtensionContext;

/**
 * Helper to check if context is WebPartContext
 */
const isWebPartContext = (context: SPFxContext): context is WebPartContext => {
	return 'domElement' in context;
};

/**
 * Helper to extract common properties from different context types
 */
const extractContextProperties = (context: SPFxContext) => {
	const baseContext = context as BaseComponentContext;

	return {
		pageContext: baseContext.pageContext,
		spHttpClient: baseContext.spHttpClient,
		serviceScope: baseContext.serviceScope,
		manifest: baseContext.manifest,
	};
};

/**
 * Comprehensive SPContext with all caching strategies and useful properties
 */
export interface ISPContext {
	// PnP SP instances with different caching strategies
	/** No caching - always fresh data */
	sp: SPFI;
	/** Standard caching - 20 minutes */
	spCached: SPFI;
	/** Minimal caching - 5 minutes */
	spMinCached: SPFI;
	/** Day caching - 24 hours */
	spDayCache: SPFI;
	/** Always cache - 7 days (for very static data) */
	spAlwaysCache: SPFI;
	/** Pessimistic cache - 1 minute (for dynamic data that might change) */
	spPessimisticCache: SPFI;

	// PnP Graph instances with different caching strategies
	/** No caching - always fresh data */
	graph: GraphFI;
	/** Standard caching - 20 minutes */
	graphCached: GraphFI;
	/** Minimal caching - 5 minutes */
	graphMinCached: GraphFI;
	/** Day caching - 24 hours */
	graphDayCache: GraphFI;
	/** Always cache - 7 days (for very static data) */
	graphAlwaysCache: GraphFI;
	/** Pessimistic cache - 1 minute (for dynamic data that might change) */
	graphPessimisticCache: GraphFI;

	// SPFx Context and useful properties
	/** Original SPFx context (WebPart or Extension context) */
	context: SPFxContext;
	/** Page context for accessing page-level information */
	pageContext: PageContext;
	/** SPHttpClient for making REST API calls */
	spHttpClient: SPHttpClient;
	/** Context type: 'WebPart' or 'Extension' */
	contextType: 'WebPart' | 'Extension';

	// URL helpers
	/** Current web relative URL (e.g., /sites/mysite/subweb) */
	webRelativeUrl: string;
	/** Current web absolute URL (e.g., https://tenant.sharepoint.com/sites/mysite/subweb) */
	webAbsoluteUrl: string;
	/** Site collection absolute URL (e.g., https://tenant.sharepoint.com/sites/mysite) */
	siteAbsoluteUrl: string;
	/** Current web server relative URL (e.g., /sites/mysite/subweb) */
	webServerRelativeUrl: string;
	/** Site collection server relative URL (e.g., /sites/mysite) */
	siteServerRelativeUrl: string;

	// User information
	/** Current user display name */
	currentUserDisplayName: string;
	/** Current user login name */
	currentUserLoginName: string;
	/** Current user email */
	currentUserEmail: string;

	// Environment information
	/** Is running in debug mode */
	isDebug: boolean;
	/** Is running in local workbench */
	isLocalWorkbench: boolean;
	/** Is running in hosted workbench */
	isHostedWorkbench: boolean;
	/** SPFx environment type */
	environmentType: string;

	// Utility methods
	/** Clear cache by type and duration */
	clearCache: (
		type?: 'sp' | 'graph' | 'all',
		duration?: 'PESSIMISTIC' | 'MINIMAL' | 'STANDARD' | 'DAY' | 'ALWAYS' | 'all'
	) => void;
	/** Get cache statistics */
	getCacheStats: () => { session: { [key: string]: number }; local: { [key: string]: number } };
	/** Check if user has specific permission */
	hasPermission: (permission: any) => boolean;
}

/**
 * Cache duration constants (in milliseconds)
 */
const CACHE_DURATIONS = {
	PESSIMISTIC: 1 * 60 * 1000, // 1 minute
	MINIMAL: 5 * 60 * 1000, // 5 minutes
	STANDARD: 20 * 60 * 1000, // 20 minutes
	DAY: 24 * 60 * 60 * 1000, // 24 hours
	ALWAYS: 7 * 24 * 60 * 60 * 1000, // 7 days
} as const;

/**
 * Create caching options for different strategies
 */
const createCacheOptions = (
	duration: number,
	keyPrefix: string,
	store: 'session' | 'local' = 'session'
): ICachingProps => ({
	store,
	keyFactory: (url: string) => `${keyPrefix}-${url}`,
	expireFunc: () => new Date(Date.now() + duration),
});

/**
 * Clear cache by type and duration
 */
const clearCacheByType = (
	type: 'sp' | 'graph' | 'all' = 'all',
	duration: keyof typeof CACHE_DURATIONS | 'all' = 'all'
): void => {
	try {
		const prefixes = type === 'all' ? ['sp', 'graph'] : [type];
		const durations =
			duration === 'all'
				? ['std', 'min', 'day', 'always', 'pess']
				: [
						duration === 'STANDARD'
							? 'std'
							: duration === 'MINIMAL'
							? 'min'
							: duration === 'DAY'
							? 'day'
							: duration === 'ALWAYS'
							? 'always'
							: 'pess',
				  ];

		// Clear session storage
		const sessionKeys = Object.keys(sessionStorage);
		sessionKeys.forEach((key) => {
			prefixes.forEach((prefix) => {
				durations.forEach((dur) => {
					if (key.includes(`${prefix}-${dur}`)) {
						sessionStorage.removeItem(key);
					}
				});
			});
		});

		// Clear local storage
		const localKeys = Object.keys(localStorage);
		localKeys.forEach((key) => {
			prefixes.forEach((prefix) => {
				durations.forEach((dur) => {
					if (key.includes(`${prefix}-${dur}`)) {
						localStorage.removeItem(key);
					}
				});
			});
		});
	} catch (error) {
		console.warn('Unable to clear cache:', error);
	}
};

/**
 * Get cache statistics for debugging
 */
const getCacheStatsInternal = (): {
	session: { [key: string]: number };
	local: { [key: string]: number };
} => {
	const stats = {
		session: {} as { [key: string]: number },
		local: {} as { [key: string]: number },
	};

	try {
		// Count session storage cache entries
		Object.keys(sessionStorage).forEach((key) => {
			if (key.includes('sp-') || key.includes('graph-')) {
				const prefix = key.split('-')[0] + '-' + key.split('-')[1];
				stats.session[prefix] = (stats.session[prefix] || 0) + 1;
			}
		});

		// Count local storage cache entries
		Object.keys(localStorage).forEach((key) => {
			if (key.includes('sp-') || key.includes('graph-')) {
				const prefix = key.split('-')[0] + '-' + key.split('-')[1];
				stats.local[prefix] = (stats.local[prefix] || 0) + 1;
			}
		});
	} catch (error) {
		console.warn('Unable to get cache stats:', error);
	}

	return stats;
};

/**
 * Global SPContext instance (singleton)
 */
let globalSPContext: ISPContext | null = null;

/**
 * Initialize the global SPContext (call this from your web part or customizer onInit)
 * @param context SPFx context (WebPart or Extension context)
 * @param logLevel Optional log level (default: Warning)
 */
export const initializeSPContext = (
	context: SPFxContext,
	logLevel: LogLevel = LogLevel.Warning
): ISPContext => {
	// Extract common properties that are available in all context types
	const contextProps = extractContextProperties(context);
	const { pageContext, spHttpClient } = contextProps;

	// Determine context type
	const contextType = isWebPartContext(context) ? 'WebPart' : 'Extension';

	// Base PnP configuration
	const baseConfig = [SPFx(context), PnPLogging(logLevel)];

	// Create SP instances with different caching strategies
	const sp = spfi().using(...baseConfig);
	const spCached = spfi().using(
		...baseConfig,
		Caching(createCacheOptions(CACHE_DURATIONS.STANDARD, 'sp-std', 'session'))
	);
	const spMinCached = spfi().using(
		...baseConfig,
		Caching(createCacheOptions(CACHE_DURATIONS.MINIMAL, 'sp-min', 'session'))
	);
	const spDayCache = spfi().using(
		...baseConfig,
		Caching(createCacheOptions(CACHE_DURATIONS.DAY, 'sp-day', 'local'))
	);
	const spAlwaysCache = spfi().using(
		...baseConfig,
		Caching(createCacheOptions(CACHE_DURATIONS.ALWAYS, 'sp-always', 'local'))
	);
	const spPessimisticCache = spfi().using(
		...baseConfig,
		Caching(createCacheOptions(CACHE_DURATIONS.PESSIMISTIC, 'sp-pess', 'session'))
	);

	// Create Graph instances with different caching strategies
	const graph = graphfi().using(...baseConfig);
	const graphCached = graphfi().using(
		...baseConfig,
		Caching(createCacheOptions(CACHE_DURATIONS.STANDARD, 'graph-std', 'session'))
	);
	const graphMinCached = graphfi().using(
		...baseConfig,
		Caching(createCacheOptions(CACHE_DURATIONS.MINIMAL, 'graph-min', 'session'))
	);
	const graphDayCache = graphfi().using(
		...baseConfig,
		Caching(createCacheOptions(CACHE_DURATIONS.DAY, 'graph-day', 'local'))
	);
	const graphAlwaysCache = graphfi().using(
		...baseConfig,
		Caching(createCacheOptions(CACHE_DURATIONS.ALWAYS, 'graph-always', 'local'))
	);
	const graphPessimisticCache = graphfi().using(
		...baseConfig,
		Caching(createCacheOptions(CACHE_DURATIONS.PESSIMISTIC, 'graph-pess', 'session'))
	);

	// Helper method to check permissions
	const hasPermission = (permission: any): boolean => {
		try {
			return pageContext.web.permissions.hasPermission(permission);
		} catch (error) {
			console.warn('Unable to check permission:', error);
			return false;
		}
	};

	// Safe environment detection - these properties may not exist in all SPFx versions
	const getEnvironmentInfo = () => {
		try {
			// Try to access diagnostics and environment properties safely
			const diagnostics = (pageContext as any).diagnostics;
			const environment = (pageContext as any).environment;

			return {
				isDebug: diagnostics?.isDebugMode || false,
				environmentType: environment?.type?.toString() || 'Unknown',
			};
		} catch (error) {
			// Fallback for older SPFx versions or when properties don't exist
			return {
				isDebug: false,
				environmentType: 'Unknown',
			};
		}
	};

	const envInfo = getEnvironmentInfo();

	// Create the comprehensive SPContext
	globalSPContext = {
		// PnP instances
		sp,
		spCached,
		spMinCached,
		spDayCache,
		spAlwaysCache,
		spPessimisticCache,
		graph,
		graphCached,
		graphMinCached,
		graphDayCache,
		graphAlwaysCache,
		graphPessimisticCache,

		// Contexts
		context,
		pageContext,
		spHttpClient,
		contextType,

		// URL properties
		webRelativeUrl: pageContext.web.absoluteUrl.replace(
			pageContext.web.absoluteUrl.split('/').slice(0, 3).join('/'),
			''
		),
		webAbsoluteUrl: pageContext.web.absoluteUrl,
		siteAbsoluteUrl: pageContext.site.absoluteUrl,
		webServerRelativeUrl: pageContext.web.serverRelativeUrl,
		siteServerRelativeUrl: pageContext.site.serverRelativeUrl,

		// User information
		currentUserDisplayName: pageContext.user.displayName,
		currentUserLoginName: pageContext.user.loginName,
		currentUserEmail: pageContext.user.email,

		// Environment information (with safe fallbacks)
		isDebug: envInfo.isDebug,
		isLocalWorkbench: pageContext.web.absoluteUrl.includes('workbench.html'),
		isHostedWorkbench: pageContext.web.absoluteUrl.includes('_layouts/15/workbench.aspx'),
		environmentType: envInfo.environmentType,

		// Utility methods
		clearCache: clearCacheByType,
		getCacheStats: getCacheStatsInternal,
		hasPermission,
	};

	return globalSPContext;
};

/**
 * Get the initialized SPContext (call this from any component)
 * @returns The global SPContext instance
 * @throws Error if SPContext has not been initialized
 */
export const getSPContext = (): ISPContext => {
	if (!globalSPContext) {
		throw new Error(
			'SPContext has not been initialized. Call initializeSPContext() first from your web part or customizer onInit method.'
		);
	}
	return globalSPContext;
};

/**
 * Check if SPContext has been initialized
 * @returns True if SPContext is initialized, false otherwise
 */
export const isSPContextInitialized = (): boolean => {
	return globalSPContext !== null;
};

/**
 * Reset the SPContext (useful for testing or reinitializing)
 */
export const resetSPContext = (): void => {
	globalSPContext = null;
};

// Additional utility functions
/**
 * Get site collection URL from context
 */
export const getSiteCollectionUrl = (context?: SPFxContext): string => {
	if (context) {
		const contextProps = extractContextProperties(context);
		return contextProps.pageContext.site.absoluteUrl;
	}
	return getSPContext().siteAbsoluteUrl;
};

/**
 * Get current web URL from context
 */
export const getCurrentWebUrl = (context?: SPFxContext): string => {
	if (context) {
		const contextProps = extractContextProperties(context);
		return contextProps.pageContext.web.absoluteUrl;
	}
	return getSPContext().webAbsoluteUrl;
};

/**
 * Get current user login name
 */
export const getCurrentUserLoginName = (context?: SPFxContext): string => {
	if (context) {
		const contextProps = extractContextProperties(context);
		return contextProps.pageContext.user.loginName;
	}
	return getSPContext().currentUserLoginName;
};

/**
 * Get current user display name
 */
export const getCurrentUserDisplayName = (context?: SPFxContext): string => {
	if (context) {
		const contextProps = extractContextProperties(context);
		return contextProps.pageContext.user.displayName;
	}
	return getSPContext().currentUserDisplayName;
};

/**
 * Get current user email
 */
export const getCurrentUserEmail = (context?: SPFxContext): string => {
	if (context) {
		const contextProps = extractContextProperties(context);
		return contextProps.pageContext.user.email;
	}
	return getSPContext().currentUserEmail;
};

/**
 * Check if current user has specific permission
 */
export const hasPermission = (permission: any, context?: SPFxContext): boolean => {
	try {
		if (context) {
			const contextProps = extractContextProperties(context);
			return contextProps.pageContext.web.permissions.hasPermission(permission);
		}
		return getSPContext().hasPermission(permission);
	} catch (error) {
		console.warn('Unable to check permission:', error);
		return false;
	}
};

/**
 * Get list server relative URL
 */
export const getListServerRelativeUrl = (listTitle: string, context?: SPFxContext): string => {
	const webServerRelativeUrl = context
		? extractContextProperties(context).pageContext.web.serverRelativeUrl
		: getSPContext().webServerRelativeUrl;
	return `${webServerRelativeUrl}/Lists/${listTitle}`;
};
