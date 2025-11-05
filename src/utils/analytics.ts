export type AnalyticsInitOptions = { endpoint: string; enabled?: boolean };

let analyticsEndpoint = '';
let analyticsEnabled = true;
let anonymousId = '';
let isFirstVisit = false;
let sessionStartTime = Date.now();

const STORAGE_KEY = 'ibm_table_plugin_anon_id';
const FIRST_VISIT_KEY = 'ibm_table_plugin_first_visit';
const LAST_VISIT_KEY = 'ibm_table_plugin_last_visit';
const SESSION_COUNT_KEY = 'ibm_table_plugin_session_count';

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

function getFirstVisitTimestamp(): number {
	try {
		const existing = localStorage.getItem(FIRST_VISIT_KEY);
		if (existing) {
			return parseInt(existing, 10);
		}
		const now = Date.now();
		localStorage.setItem(FIRST_VISIT_KEY, String(now));
		isFirstVisit = true;
		return now;
	} catch {
		return Date.now();
	}
}

function getLastVisitTimestamp(): number | null {
	try {
		const lastVisit = localStorage.getItem(LAST_VISIT_KEY);
		return lastVisit ? parseInt(lastVisit, 10) : null;
	} catch {
		return null;
	}
}

function updateLastVisit() {
	try {
		localStorage.setItem(LAST_VISIT_KEY, String(Date.now()));
	} catch {
		// Ignore
	}
}

function incrementSessionCount(): number {
	try {
		const countStr = localStorage.getItem(SESSION_COUNT_KEY);
		const count = countStr ? parseInt(countStr, 10) + 1 : 1;
		localStorage.setItem(SESSION_COUNT_KEY, String(count));
		return count;
	} catch {
		return 1;
	}
}

export function init(options: AnalyticsInitOptions) {
	analyticsEndpoint = options.endpoint;
	analyticsEnabled = options.enabled !== false;
	anonymousId = getAnonymousId();
	sessionStartTime = Date.now();
	
	const firstVisit = getFirstVisitTimestamp();
	const lastVisit = getLastVisitTimestamp();
	const sessionCount = incrementSessionCount();
	const daysSinceLastVisit = lastVisit ? Math.floor((Date.now() - lastVisit) / (1000 * 60 * 60 * 24)) : null;
	
	updateLastVisit();
	
	// eslint-disable-next-line no-console
	console.log('[analytics] initialized', { 
		endpoint: analyticsEndpoint, 
		enabled: analyticsEnabled,
		isFirstVisit,
		sessionCount 
	});
	
	// Track session start with user context
	track('session_start', {
		is_first_visit: isFirstVisit,
		session_count: sessionCount,
		days_since_last_visit: daysSinceLastVisit,
		user_type: isFirstVisit ? 'new' : 'returning'
	});
}

export async function track(event: string, props: Record<string, any> = {}) {
	if (!analyticsEnabled || !analyticsEndpoint) return;
	
	// Add user context to all events
	const enrichedProps = {
		...props,
		user_id: anonymousId,
		// Session info
		session_duration_seconds: Math.floor((Date.now() - sessionStartTime) / 1000),
		// User type (will be set by init for session_start, but can be overridden)
		...(!props.user_type && { user_type: isFirstVisit ? 'new' : 'returning' }),
	};
	
	const payload = {
		event,
		anonId: anonymousId,
		props: sanitize(enrichedProps),
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

export function trackSessionEnd() {
	const sessionDuration = Math.floor((Date.now() - sessionStartTime) / 1000);
	track('session_end', {
		session_duration_seconds: sessionDuration,
	});
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
