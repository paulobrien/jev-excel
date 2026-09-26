/*
 * Jev for Excel — Cloudflare Worker.
 *
 * Hosts the whole add-in from one origin (see wrangler.toml):
 *  - Static assets (the built ./dist folder) are served by Cloudflare directly.
 *  - /v1/systemone and /v1/models are proxied to TypeSafe. api.typesafe.ai rejects
 *    browser (CORS) requests, so the add-in calls its own origin instead.
 *  - /manifest.xml is the development manifest with every https://localhost:3000
 *    URL replaced by this Worker's origin, so it's right on any domain.
 *
 * The proxy holds no key. Each caller's own "Authorization: Bearer <TypeSafe key>"
 * header is passed through unchanged, so usage is billed to the caller's account.
 * This code logs and stores nothing.
 */
const UPSTREAM = "https://api.typesafe.ai";
const API_ROUTES = { "/v1/systemone": "POST", "/v1/models": "GET" };
const PASS_HEADERS = ["Content-Type", "Retry-After", "retry-after-ms", "x-typesafe-request-id"];
const MAX_BODY_BYTES = 1_000_000;
const DEV_ORIGIN = "https://localhost:3000";

function json(status, message) {
  return new Response(JSON.stringify({ error: { message } }), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

async function proxy(request, method, pathname) {
  if (request.method !== method) return json(405, `Use ${method}`);
  const auth = request.headers.get("Authorization");
  if (!auth) return json(401, "Missing API key: set your TypeSafe key in the Jev pane");

  let body;
  if (method === "POST") {
    if (Number(request.headers.get("Content-Length")) > MAX_BODY_BYTES) return json(413, "Request too large");
    body = await request.arrayBuffer();
    if (body.byteLength > MAX_BODY_BYTES) return json(413, "Request too large");
  }

  let upstream;
  try {
    upstream = await fetch(UPSTREAM + pathname, {
      method,
      headers: {
        Authorization: auth,
        Accept: "application/json",
        ...(body !== undefined ? { "Content-Type": "application/json" } : {}),
      },
      body,
    });
  } catch (_) {
    return json(502, `Could not reach ${UPSTREAM}`);
  }

  const headers = new Headers({ "Cache-Control": "no-store" });
  for (const name of PASS_HEADERS) {
    const value = upstream.headers.get(name);
    if (value) headers.set(name, value);
  }
  return new Response(upstream.body, { status: upstream.status, headers });
}

async function manifest(request, env) {
  const template = await env.ASSETS.fetch(request);
  if (!template.ok) return template;
  const { origin } = new URL(request.url);
  const xml = (await template.text()).split(DEV_ORIGIN).join(origin);
  return new Response(xml, { headers: { "Content-Type": "application/xml; charset=utf-8", "Cache-Control": "no-cache" } });
}

export default {
  async fetch(request, env) {
    const { pathname } = new URL(request.url);
    const method = API_ROUTES[pathname];
    if (method) return proxy(request, method, pathname);
    if (pathname === "/manifest.xml") return manifest(request, env);
    return env.ASSETS.fetch(request);
  },
};
