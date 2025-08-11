// utils/assert.ts
export function assert(cond: any, msg = 'Assertion failed'): asserts cond {
	if (!cond) throw new Error(msg);
}

export function coerceBoolean(v: any, def = false) {
	if (typeof v === 'boolean') return v;
	if (typeof v === 'string') return ['true', '1', 'yes', 'y', 'on'].includes(v.toLowerCase());
	if (typeof v === 'number') return v !== 0;
	return def;
}

export function asNumber(v: any, def = 0) {
	const n = Number(v);
	return Number.isFinite(n) ? n : def;
}

export async function sleep(ms: number) {
	return new Promise((res) => setTimeout(res, ms));
}
