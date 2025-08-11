// Lightweight helpers used across the Upload + Edit solution

import { NormalizedError, TelemetryTimer, FileProgress } from './types';

/* ---------------------------------- Errors --------------------------------- */

export function normalizeError(e: unknown, fallback = 'Something went wrong'): NormalizedError {
	if (!e) return { message: fallback };
	if (typeof e === 'string') return { message: e };
	if (e instanceof Error)
		return { message: e.message || fallback, cause: e, code: (e as any).code };
	try {
		const msg =
			(e as any)?.message ||
			(e as any)?.error?.message ||
			(e as any)?.statusText ||
			JSON.stringify(e);
		return { message: msg || fallback, cause: e };
	} catch {
		return { message: fallback };
	}
}

/* ------------------------------ Telemetry timer ----------------------------- */

export function createTelemetryTimer(): TelemetryTimer {
	const marks = new Map<string, number>();
	return {
		start(label: string) {
			marks.set(label, performance.now());
		},
		end(label: string) {
			const start = marks.get(label);
			const dur = start ? performance.now() - start : 0;
			marks.delete(label);
			return dur;
		},
	};
}

/* --------------------------------- Strings --------------------------------- */

export function formatBytes(bytes: number): string {
	if (!Number.isFinite(bytes) || bytes < 0) return `${bytes}`;
	const units = ['B', 'KB', 'MB', 'GB', 'TB'];
	let i = 0;
	let v = bytes;
	while (v >= 1024 && i < units.length - 1) {
		v /= 1024;
		i++;
	}
	return `${v.toFixed(v < 10 && i > 0 ? 1 : 0)} ${units[i]}`;
}

/* --------------------------------- Paths ----------------------------------- */

export function safeJoinPath(base: string, sub?: string): string {
	if (!sub) return trimTrailingSlash(base);
	return `${trimTrailingSlash(base)}/${trimLeadingSlash(sub)}`;
}

export function trimLeadingSlash(p: string): string {
	return p.replace(/^\/+/, '');
}

export function trimTrailingSlash(p: string): string {
	return p.replace(/\/+$/, '');
}

export function encodePathSegments(path: string): string {
	// Encode each segment but keep slashes
	return path.split('/').map(encodeURIComponent).join('/');
}

/* ---------------------------- Filenames / suffix ---------------------------- */

export function getNameAndExt(fileName: string): { name: string; ext: string } {
	const i = fileName.lastIndexOf('.');
	if (i <= 0 || i === fileName.length - 1) return { name: fileName, ext: '' };
	return { name: fileName.slice(0, i), ext: fileName.slice(i) };
}

/**
 * Produce "Report (1).docx", "Report (2).docx", etc.
 * Provide attempt starting at 1.
 */
export function withNumericSuffix(fileName: string, attempt: number): string {
	const { name, ext } = getNameAndExt(fileName);
	return `${name} (${attempt})${ext}`;
}

/* ------------------------------ Progress utils ----------------------------- */

/** Average progress across files (simple mean of percents) */
export function overallPercent(progress: FileProgress[]): number {
	if (!progress?.length) return 0;
	const sum = progress.reduce((acc, p) => acc + (p.percent || 0), 0);
	return Math.round(sum / progress.length);
}

/* --------------------------------- Throttle -------------------------------- */

export function throttle<T extends (...args: any[]) => void>(fn: T, ms = 100): T {
	let last = 0;
	let timer: any = null;
	let lastArgs: any[] | null = null;

	const invoke = () => {
		last = Date.now();
		timer = null;
		fn.apply(null, lastArgs as any[]);
		lastArgs = null;
	};

	const wrapped = ((...args: any[]) => {
		lastArgs = args;
		const now = Date.now();
		const remaining = ms - (now - last);
		if (remaining <= 0) {
			if (timer) {
				clearTimeout(timer);
				timer = null;
			}
			invoke();
		} else if (!timer) {
			timer = setTimeout(invoke, remaining);
		}
	}) as T;

	return wrapped;
}

/* --------------------------------- Debounce -------------------------------- */

export function debounce<T extends (...args: any[]) => void>(fn: T, ms = 150): T {
	let timer: any = null;
	const wrapped = ((...args: any[]) => {
		if (timer) clearTimeout(timer);
		timer = setTimeout(() => fn.apply(null, args), ms);
	}) as T;
	return wrapped;
}

/* --------------------------------- Guards ---------------------------------- */

export function assert<T>(cond: T, message = 'Assertion failed'): asserts cond {
	if (!cond) throw new Error(message);
}

export function isNonEmptyString(v: unknown): v is string {
	return typeof v === 'string' && v.trim().length > 0;
}
