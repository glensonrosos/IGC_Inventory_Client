import axios from 'axios';

// Configure API base via Vite env. Fallback to '/api' (proxied in dev).
const apiBase = (import.meta as any)?.env?.VITE_API_BASE || '/api';
const api = axios.create({ baseURL: apiBase });

// Prevent multiple concurrent redirects on auth expiry
let __authRedirecting = false;

function redirectToLoginOnce() {
  if (typeof window === 'undefined') return;
  if (__authRedirecting) return;
  __authRedirecting = true;
  try {
    localStorage.removeItem('token');
  } catch {}
  const path = window.location.pathname || '/';
  const search = window.location.search || '';
  const next = encodeURIComponent(`${path}${search}`);
  // Avoid redirect loop if already on /login
  if (path.startsWith('/login')) {
    __authRedirecting = false; // allow future redirects after manual action
    return;
  }
  // Use assign to create a full navigation (not SPA push) to ensure clean state
  window.location.assign(`/login?next=${next}`);
}

api.interceptors.request.use((config) => {
  const token = localStorage.getItem('token');
  if (token) {
    if (!config.headers) config.headers = {} as any;
    (config.headers as any)['Authorization'] = `Bearer ${token}`;
  }
  (config as any)._ts = Date.now();
  const debug = (import.meta as any)?.env?.DEV || localStorage.getItem('DEBUG_API') === '1';
  if (debug) {
    const method = (config.method || 'get').toUpperCase();
    const url = `${config.baseURL || ''}${config.url || ''}`;
    const info = { params: config.params, data: config.data };
    try { console.log('[API] ->', method, url, info); } catch {}
  }
  return config;
});

api.interceptors.response.use(
  (response) => {
    const debug = (import.meta as any)?.env?.DEV || localStorage.getItem('DEBUG_API') === '1';
    if (debug) {
      const cfg: any = response.config || {};
      const started = Number(cfg._ts || 0);
      const ms = started ? Date.now() - started : undefined;
      const method = (cfg.method || 'get').toUpperCase();
      const url = `${cfg.baseURL || ''}${cfg.url || ''}`;
      const status = response.status;
      const info = { params: cfg.params, data: cfg.data };
      try { console.log('[API] <-', method, url, status, ms !== undefined ? `${ms}ms` : '', info); } catch {}
    }
    return response;
  },
  (error) => {
    const debug = (import.meta as any)?.env?.DEV || localStorage.getItem('DEBUG_API') === '1';
    if (debug) {
      const cfg: any = error?.config || {};
      const started = Number(cfg._ts || 0);
      const ms = started ? Date.now() - started : undefined;
      const method = (cfg.method || 'get').toUpperCase();
      const url = `${cfg.baseURL || ''}${cfg.url || ''}`;
      const status = error?.response?.status;
      const info = { params: cfg.params, data: cfg.data };
      const msg = error?.message || 'Request failed';
      try { console.log('[API] x ', method, url, status ?? '-', ms !== undefined ? `${ms}ms` : '', msg, info); } catch {}
    }

    // Handle auth expiration: redirect to login on 401/403
    try {
      const status = error?.response?.status;
      const cfgUrl: string = (error?.config?.url || '').toString();
      const isAuthEndpoint = /\/auth\b|\/login\b/i.test(cfgUrl);
      if ((status === 401 || status === 403) && !isAuthEndpoint) {
        redirectToLoginOnce();
      }
    } catch {}
    return Promise.reject(error);
  }
);

export default api;
