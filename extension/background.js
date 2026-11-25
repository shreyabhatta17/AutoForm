// Where the AutoForm API runs. Change DEFAULT_API to your deployed backend URL
// (and add the same origin to host_permissions in manifest.json).
// For local development you can override it without editing code, from the
// extension's devtools console:  chrome.storage.local.set({ apiBase: 'http://127.0.0.1:8000' })
const DEFAULT_API = 'https://autoform-shreya.onrender.com';

async function apiBase() {
  const { apiBase } = await chrome.storage.local.get('apiBase');
  return String(apiBase || DEFAULT_API).replace(/\/+$/, '');
}

// One anonymous account per extension install (per backend URL). The key is
// created on first use and kept in extension storage.
async function apiKey(base, { renew = false } = {}) {
  const { apiKeys = {} } = await chrome.storage.local.get('apiKeys');
  if (apiKeys[base] && !renew) return apiKeys[base];
  const response = await fetch(`${base}/auth/register`, { method: 'POST' });
  if (!response.ok) throw new Error((await response.json().catch(() => ({}))).detail || 'Could not register with the backend.');
  const { api_key } = await response.json();
  apiKeys[base] = api_key;
  await chrome.storage.local.set({ apiKeys });
  return api_key;
}

async function request(path, options = {}, retried = false) {
  const base = await apiBase();
  const key = await apiKey(base);
  const response = await fetch(`${base}${path}`, {
    ...options,
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${key}`, ...(options.headers || {}) }
  });
  if (response.status === 401 && !retried) {
    // The server no longer knows this key (e.g. account deleted): start a fresh account.
    await apiKey(base, { renew: true });
    return request(path, options, true);
  }
  if (!response.ok) throw new Error((await response.json().catch(() => ({}))).detail || `Backend error (${response.status})`);
  return response.json();
}

async function health() {
  const base = await apiBase();
  const response = await fetch(`${base}/health`);
  if (!response.ok) throw new Error('offline');
  return response.json();
}

async function openDashboard() {
  const base = await apiBase();
  const key = await apiKey(base);
  // The key travels in the URL fragment, which browsers never send to the server.
  await chrome.tabs.create({ url: `${base}/#key=${encodeURIComponent(key)}` });
  return { ok: true };
}

const handlers = {
  health: () => health().catch(() => ({ status: 'offline' })),
  classify: m => request('/fields/classify', { method: 'POST', body: JSON.stringify({ fields: m.fields }) }),
  generate: m => request('/answers/generate', { method: 'POST', body: JSON.stringify({ question: m.question, length: m.length }) }),
  'save-answer': m => request('/answers/save', { method: 'POST', body: JSON.stringify({ question: m.question, answer: m.answer }) }),
  'open-dashboard': () => openDashboard()
};

chrome.runtime.onMessage.addListener((message, _sender, reply) => {
  const handler = handlers[message?.type];
  if (!handler) return false;
  handler(message).then(reply).catch(error => reply({ error: error.message || String(error) }));
  return true;
});
