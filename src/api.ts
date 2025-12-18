import axios from 'axios';

// Configure API base via Vite env. Fallback to '/api' (proxied in dev).
const apiBase = (import.meta as any)?.env?.VITE_API_BASE || '/api';
const api = axios.create({ baseURL: apiBase });

api.interceptors.request.use((config) => {
  const token = localStorage.getItem('token');
  if (token) config.headers = { ...config.headers, Authorization: `Bearer ${token}` };
  return config;
});

export default api;
