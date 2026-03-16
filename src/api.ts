import axios from 'axios';

// Configure API base via Vite env. Fallback to '/api' (proxied in dev).
const apiBase = (import.meta as any)?.env?.VITE_API_BASE || '/api';
const api = axios.create({ baseURL: apiBase });

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
    return Promise.reject(error);
  }
);

export default api;
