export type AnalyticsInitOptions = { endpoint: string; enabled?: boolean };

let analyticsEndpoint = '';
let analyticsEnabled = true;
let anonymousId = '';

const STORAGE_KEY = 'ibm_table_plugin_anon_id';

function getAnonymousId(): string {
	try {
		const existing = localStorage.getItem(STORAGE_KEY);
		if (existing) return existing;
		const id = 'anon_' + Math.random().toString(36).slice(2);
		localStorage.setItem(STORAGE_KEY, id);
		return id;
	} catch {
		return 'anon_' + Math.random().toString(36).slice(2);
	}
}

export function init(options: AnalyticsInitOptions) {
	analyticsEndpoint = options.endpoint;
	analyticsEnabled = options.enabled !== false;
	anonymousId = getAnonymousId();
	// eslint-disable-next-line no-console
	console.log('[analytics] initialized', { endpoint: analyticsEndpoint, enabled: analyticsEnabled });
}

export async function track(event: string, props: Record<string, any> = {}) {
	if (!analyticsEnabled || !analyticsEndpoint) return;
	const payload = {
		event,
		anonId: anonymousId,
		props: sanitize(props),
		ts: Date.now(),
	};
	try {
		await fetch(analyticsEndpoint, {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify(payload),
		});
	} catch (e) {
		// eslint-disable-next-line no-console
		console.warn('[analytics] send failed', e);
	}
}

function sanitize(obj: Record<string, any>) {
	const clean: Record<string, any> = {};
	for (const [k, v] of Object.entries(obj || {})) {
		if (v === null || v === undefined) continue;
		if (typeof v === 'string' && v.length > 1000) {
			clean[k] = v.slice(0, 1000);
			continue;
		}
		clean[k] = v;
	}
	return clean;
}
