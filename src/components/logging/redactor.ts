// logging/redactor.ts
const SECRET_KEYS = [
	'authorization',
	'cookie',
	'set-cookie',
	'bearer',
	'token',
	'apikey',
	'clientsecret',
	'password',
];

export function redactDeep(input: any): any {
	try {
		return _redact(input, new WeakSet());
	} catch {
		return input;
	}
}

function _redact(v: any, seen: WeakSet<any>): any {
	if (v === null || typeof v !== 'object') return v;
	if (seen.has(v)) return '[Circular]';
	seen.add(v);

	if (Array.isArray(v)) return v.map((x) => _redact(x, seen));

	const out: any = {};
	for (const k of Object.keys(v)) {
		if (SECRET_KEYS.some((s) => k.toLowerCase().includes(s))) out[k] = '[REDACTED]';
		else out[k] = _redact(v[k], seen);
	}
	return out;
}
