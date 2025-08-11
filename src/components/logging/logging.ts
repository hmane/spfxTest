// logging/logging.ts
import { Logger as PnPLogger, ILogEntry, LogLevel } from '@pnp/logging';
import { Log } from '@microsoft/sp-core-library';
import type { ApplicationInsights } from '@microsoft/applicationinsights-web';
import { redactDeep } from './redactor';

export type LogCategory = string;

export interface LogEvent {
	level: LogLevel;
	message: string;
	category?: LogCategory;
	data?: any;
	time?: number;
}

export interface LoggerOptions {
	level: LogLevel;
	defaultCategory?: LogCategory;
	enrich?: () => Record<string, any>;
	redactor?: (o: any) => any;
}

export interface LogSink {
	log(e: LogEvent & { enriched?: any }): void;
}

/** Heuristic parser for SharePoint/Graph error payloads. */
function extractSpErrorFields(from: any): {
	spMessage?: string;
	spCode?: string;
	raw?: any;
} {
	if (!from) return {};
	// Try common shapes:
	// Graph: { error: { code, message, innerError } }
	if (from.error) {
		const e = from.error;
		const spMessage = typeof e.message === 'string' ? e.message : e.message?.value;
		const spCode = e.code;
		return { spMessage, spCode, raw: from };
	}
	// OData v3/v4: { "odata.error": { message: { value }, code } }
	if (from['odata.error']) {
		const e = from['odata.error'];
		const spMessage = e?.message?.value;
		const spCode = e?.code;
		return { spMessage, spCode, raw: from };
	}
	// Sometimes it is { message: { value } }
	if (from.message?.value) {
		return { spMessage: from.message.value, raw: from };
	}
	// Fallback: message string on root
	if (typeof from.message === 'string') {
		return { spMessage: from.message, raw: from };
	}
	return { raw: from };
}

/** Normalize unknown errors into a consistent structure for logging. */
function normalizeError(err: unknown): {
	message: string;
	name?: string;
	status?: number;
	statusText?: string;
	stack?: string;
	spMessage?: string;
	spCode?: string;
	original?: any;
} {
	// PnP's HttpRequestError has status/statusText/isHttpRequestError
	const maybe: any = err as any;
	const name = maybe?.name || (maybe?.isHttpRequestError ? 'HttpRequestError' : undefined);
	const status = maybe?.status;
	const statusText = maybe?.statusText;
	const stack = typeof maybe?.stack === 'string' ? maybe.stack : undefined;

	// Try to pull a JSON body if one was attached (some libs attach .data or .body)
	let parsed: any | undefined;
	const body = maybe?.body || maybe?.data || maybe?.responseBody;
	if (body && typeof body === 'string') {
		try {
			parsed = JSON.parse(body);
		} catch {
			/* ignore */
		}
	} else if (typeof maybe?.response === 'object' && typeof maybe?.response?.text === 'function') {
		// In some thrown shapes response is a fetch Response; we can't await here (sync),
		// so the caller should prefer passing { status, body } when possible.
	}

	const fields = extractSpErrorFields(parsed ?? maybe);
	const msg = fields.spMessage || maybe?.message || statusText || 'Unexpected error';

	return {
		message: msg,
		name,
		status,
		statusText,
		stack,
		spMessage: fields.spMessage,
		spCode: fields.spCode,
		original: err,
	};
}

export class ConsoleSink implements LogSink {
	log(e: LogEvent & { enriched?: any }) {
		const cat = e.category || 'App';
		const base = `[${cat}]`;
		const payload = e.enriched ?? e.data;
		switch (e.level) {
			case LogLevel.Error:
				console.error(base, e.message, payload ?? '');
				break;
			case LogLevel.Warning:
				console.warn(base, e.message, payload ?? '');
				break;
			case LogLevel.Info:
				console.info(base, e.message, payload ?? '');
				break;
			default:
				console.debug(base, e.message, payload ?? '');
		}
	}
}

export class DiagnosticsSink implements LogSink {
	log(e: LogEvent & { enriched?: any }) {
		const cat = e.category || 'App';
		const payload = JSON.stringify(e.enriched ?? e.data ?? {});
		if (e.level === LogLevel.Error) Log.error(cat, new Error(`${e.message} ${payload}`));
		else if (e.level === LogLevel.Warning) Log.warn(cat, `${e.message} ${payload}`);
		else Log.info(cat, `${e.message} ${payload}`);
	}
}

export class AppInsightsSink implements LogSink {
	constructor(private ai: ApplicationInsights) {}
	log(e: LogEvent & { enriched?: any }) {
		const props = e.enriched ?? e.data ?? {};
		if (e.level === LogLevel.Error) {
			this.ai.trackException({ exception: new Error(e.message), properties: props });
		} else {
			this.ai.trackTrace({
				message: e.message,
				properties: props,
				severityLevel: toSeverity(e.level),
			});
		}
	}
}

function toSeverity(l: LogLevel): 0 | 1 | 2 | 3 {
	return l === LogLevel.Error ? 3 : l === LogLevel.Warning ? 2 : l === LogLevel.Info ? 1 : 0;
}

export class LoggerFacade {
	private sinks: LogSink[] = [];
	constructor(private opts: LoggerOptions) {}

	addSink(s: LogSink) {
		this.sinks.push(s);
	}
	removeSink(s: LogSink) {
		this.sinks = this.sinks.filter((x) => x !== s);
	}

	private emit(level: LogLevel, message: string, data?: any, category?: string) {
		if (level > this.opts.level) return;
		const enriched = this.opts.enrich ? this.opts.enrich() : {};
		const safe = this.opts.redactor ? this.opts.redactor({ ...data }) : data;
		const evt: LogEvent & { enriched?: any } = {
			level,
			message,
			category: category ?? this.opts.defaultCategory,
			data,
			time: Date.now(),
			enriched: { ...enriched, ...safe },
		};
		for (const s of this.sinks) s.log(evt);
	}

	// ---------- Core levels ----------
	info(msg: string, data?: any, category?: string) {
		this.emit(LogLevel.Info, msg, data, category);
	}
	warn(msg: string, data?: any, category?: string) {
		this.emit(LogLevel.Warning, msg, data, category);
	}
	/** Error: accepts (message, data) OR (errorObject) */
	error(errOrMsg: unknown, data?: any, category?: string) {
		if (typeof errOrMsg === 'string') {
			this.emit(LogLevel.Error, errOrMsg, data, category);
		} else {
			const n = normalizeError(errOrMsg);
			// Merge caller data (if any) after normalized fields
			this.emit(LogLevel.Error, n.message, { ...n, ...(data ?? {}) }, category);
		}
	}
	/** Verbose: low-priority, chatty logging (hidden unless level >= Verbose) */
	verbose(msg: string, data?: any, category?: string) {
		this.emit(LogLevel.Verbose, msg, data, category);
	}

	// ---------- Developer-friendly sugar ----------
	/** Green checkmark success line (Info level) */
	success(msg: string, data?: any, category?: string) {
		this.info(`✅ ${msg}`, data, category);
	}
	/** ❌ Convenience wrapper at Error level; also accepts Error/unknown */
	fail(errOrMsg: unknown, data?: any, category?: string) {
		if (typeof errOrMsg === 'string') {
			this.error(`❌ ${errOrMsg}`, data, category);
		} else {
			const n = normalizeError(errOrMsg);
			this.error(`❌ ${n.message}`, { ...n, ...(data ?? {}) }, category);
		}
	}

	banner(title: string, details?: any, category?: string) {
		const line = '─'.repeat(Math.max(10, Math.min(60, title.length + 10)));
		this.info(line, undefined, category);
		this.info(`★ ${title}`, details, category);
		this.info(line, undefined, category);
	}
	step(n: number, msg: string, data?: any, category?: string) {
		this.info(`Step ${n}: ${msg}`, data, category);
	}
	kv(obj: Record<string, any>, category?: string) {
		this.info('kv', obj, category);
	}
	table(rows: any[], columns?: string[], category?: string) {
		if (console.table) console.table(rows, columns);
		else this.info('table', { rows, columns }, category);
	}
	code(label: string, codeOrObject: any, category?: string) {
		const text =
			typeof codeOrObject === 'string' ? codeOrObject : JSON.stringify(codeOrObject, null, 2);
		this.info(`${label}\n${text}`, undefined, category);
	}
	link(text: string, url: string, category?: string) {
		this.info(`${text}: ${url}`, undefined, category);
	}
	hr(category?: string) {
		this.info('—'.repeat(40), undefined, category);
	}
}

export function createLogger(opts: LoggerOptions) {
	return new LoggerFacade(opts);
}

export function bridgePnPLoggerToSinks(logger: LoggerFacade) {
	PnPLogger.subscribe({
		log: (e: ILogEntry) => {
			const data = e.data ? redactDeep(e.data) : undefined;
			// ILogEntry in v3 has no "category" field; use defaultCategory
			if (e.level === LogLevel.Error) logger.error(e.message, data);
			else if (e.level === LogLevel.Warning) logger.warn(e.message, data);
			else if (e.level === LogLevel.Info) logger.info(e.message, data);
			else logger.verbose(e.message, data);
		},
	});
}
