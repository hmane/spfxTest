// net/telemetry.ts
import { ApplicationInsights } from '@microsoft/applicationinsights-web';

export function initAppInsights(key: string, roleName = 'SPFxApp') {
	const ai = new ApplicationInsights({
		config: {
			instrumentationKey: key,
			enableAutoRouteTracking: false,
			disableAjaxTracking: true,
			disableFetchTracking: true,
			enableUnhandledPromiseRejectionTracking: true,
		},
	});
	ai.loadAppInsights();
	ai.addTelemetryInitializer((env) => {
		(env.tags as any) = env.tags || {};
		(env.tags as any)['ai.cloud.role'] = roleName;
	});
	return ai;
}
