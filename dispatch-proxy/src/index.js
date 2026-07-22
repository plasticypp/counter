// Cloudflare Worker: proxies the Dispatch Dashboard's data from Apps Script and serves the
// frontend directly from this Worker. The Apps Script /exec URL lives ONLY in env.GAS_URL
// (a secret, never sent to the browser) — the client only ever talks to this Worker's own
// domain, so the Apps Script backend is fully hidden.

import DASHBOARD_HTML from './dashboard.html';

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === '/api/data') {
      return handleData(env);
    }

    return new Response(DASHBOARD_HTML, {
      headers: { 'content-type': 'text/html; charset=utf-8' },
    });
  },
};

async function handleData(env) {
  if (!env.GAS_URL) {
    return json({ error: 'GAS_URL secret not configured' }, 500);
  }
  try {
    const res = await fetch(env.GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({ action: 'getDashboardData' }),
    });
    if (!res.ok) return json({ error: 'Upstream error: ' + res.status }, 502);
    const data = await res.json();
    return json(data);
  } catch (err) {
    return json({ error: String(err) }, 502);
  }
}

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: {
      'content-type': 'application/json',
      // Short cache: dashboard data changes as Dispatch is edited, but this avoids hammering
      // Apps Script on rapid refreshes/tab-switches.
      'cache-control': 'private, max-age=10',
    },
  });
}
