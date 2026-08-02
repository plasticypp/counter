// Cloudflare Worker: proxies the Dispatch Dashboard's data from Apps Script and serves the
// frontend directly from this Worker. The Apps Script /exec URL lives ONLY in env.GAS_URL
// (a secret, never sent to the browser) — the client only ever talks to this Worker's own
// domain, so the Apps Script backend is fully hidden.

import DASHBOARD_HTML from './dashboard.html';

// Cache keys. Declared here, above first use: `const` is not hoisted the way a function
// declaration is, so defining these below handleData threw a ReferenceError on every single
// request — the worker returned Cloudflare's own 500 page and the board went blank.
const CACHE_KEY = 'https://dispatch-proxy.internal/api/data';
const STALE_KEY = 'https://dispatch-proxy.internal/api/data-last-good';

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    if (url.pathname === '/api/data') {
      return handleData(env, ctx);
    }

    return new Response(DASHBOARD_HTML, {
      headers: {
        'content-type': 'text/html; charset=utf-8',
        // Always revalidate. The page carries its own JS, so a cached copy keeps running
        // old code after a deploy — which has repeatedly looked like a broken dashboard
        // rather than a stale one. It is a single small document; the round trip is cheap.
        'cache-control': 'no-cache, must-revalidate',
      },
    });
  },
};

async function handleData(env, ctx) {
  if (!env.GAS_URL) {
    return json({ error: 'GAS_URL secret not configured' }, 500);
  }
  // WHY THIS CACHES AT THE EDGE
  //
  // Apps Script itself is fast and reliable: measured directly, five consecutive calls took
  // 2.0-3.0s with no failures. Through this worker the same endpoint was taking anywhere up
  // to 52s and intermittently returning 404. The upstream was never the problem — the
  // request VOLUME was.
  //
  // Apps Script serialises concurrent /exec invocations per script and sheds load when they
  // pile up; that shedding is the transient 404. Every viewer polling independently, with
  // the browser cache bypassed on forced refreshes, meant each one queued behind the others.
  // A retry made it worse by tripling the pressure that caused the queue.
  //
  // Caching in the Cloudflare cache — shared across every viewer, unlike a browser cache —
  // collapses all of that into at most one upstream call per TTL, however many phones and
  // tablets are watching.
  const cache = caches.default;
  const cacheKey = new Request(CACHE_KEY, { method: 'GET' });

  const hit = await cache.match(cacheKey);
  if (hit) return hit;

  // A stale copy is kept purely as a fallback. When Apps Script throttles — which it does
  // under load, answering with an HTML error page, a 404, or a 200 carrying zero rows — the
  // honest choice is last-known-good data with its age shown, not a blank board. The
  // freshness stamp already tells the reader how old what they are seeing is.
  const stale = await cache.match(new Request(STALE_KEY));

  let data = null;
  let problem = null;

  try {
    const upstream = await fetch(env.GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({ action: 'getDashboardData' }),
    });

    const body = await upstream.text();

    if (!upstream.ok) {
      problem = 'HTTP ' + upstream.status;
    } else if (body.trimStart().startsWith('<')) {
      // Apps Script serves an HTML error page rather than JSON when it sheds load.
      problem = 'throttled by Apps Script';
    } else {
      try {
        const parsed = JSON.parse(body);
        if (parsed && Array.isArray(parsed.rows) && parsed.rows.length) data = parsed;
        else problem = 'empty response';
      } catch (e) {
        problem = 'malformed response';
      }
    }
  } catch (err) {
    problem = String(err);
  }

  if (data) {
    const res = json(data);
    // Two entries: the short-lived one every viewer shares, and a long-lived copy that
    // survives an outage so the board can fall back to it.
    ctx.waitUntil(cache.put(cacheKey, res.clone()));
    ctx.waitUntil(cache.put(new Request(STALE_KEY), staleCopy(data)));
    return res;
  }

  if (stale) {
    // .clone() because a cached Response body can only be consumed once; reading the
    // original would leave the next fallback with an unusable entry.
    const body = await stale.clone().json();
    body.stale = true;
    body.staleReason = problem;
    // Deliberately not re-cached: the next request should attempt the upstream again rather
    // than be handed this copy.
    return json(body);
  }

  return json({ error: 'Schedule unavailable (' + problem + ')' }, 502);
}

/** Long-TTL copy used only when the upstream is failing. */
function staleCopy(data) {
  return new Response(JSON.stringify(data), {
    headers: {
      'content-type': 'application/json',
      'cache-control': 'public, s-maxage=86400',
    },
  });
}

function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: {
      'content-type': 'application/json',
      // s-maxage governs the SHARED edge cache; max-age governs the browser. `private` was
      // wrong here — it forbids exactly the shared caching that protects Apps Script from
      // every viewer arriving at once.
      //
      // 45s at the edge against a 2-minute client poll means a viewer never sees data older
      // than their own refresh interval, while a dozen phones still cost one upstream call.
      'cache-control': 'public, max-age=15, s-maxage=45',
    },
  });
}
