#!/usr/bin/env node
/*
 * HTTPS server for local development, on https://localhost:3000 by default.
 *
 *  - Serves ./src using the Office dev certificates (run `npm run certs` once).
 *  - Proxies /v1/systemone and /v1/models to TypeSafe. api.typesafe.ai rejects
 *    browser (CORS) requests, so the add-in can't call it directly; when it's
 *    loaded from here, its Base URL defaults to this server instead.
 *
 * Like the Cloudflare Worker, the proxy holds no key: the caller's own
 * "Authorization: Bearer <TypeSafe key>" header is passed through unchanged.
 */
"use strict";
const https = require("node:https");
const fs = require("node:fs");
const path = require("node:path");

const PORT = Number(process.env.PORT) || 3000;
const ORIGIN = `https://localhost:${PORT}`;
const ROOT = path.resolve(__dirname, "..", "src");
const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".xml": "application/xml; charset=utf-8",
};

const UPSTREAM = "https://api.typesafe.ai";
const API_ROUTES = { "/v1/systemone": "POST", "/v1/models": "GET" };
const PASS_HEADERS = ["content-type", "retry-after", "retry-after-ms", "x-typesafe-request-id"];
const MAX_BODY_BYTES = 1_000_000;
const UPSTREAM_TIMEOUT_MS = 30_000;

class HttpError extends Error {
  constructor(status, message) {
    super(message);
    this.status = status;
  }
}

function sendJson(res, status, body) {
  res.writeHead(status, { "Content-Type": "application/json", "Cache-Control": "no-store" });
  res.end(JSON.stringify(body));
}

async function readBody(req) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > MAX_BODY_BYTES) throw new HttpError(413, "Request too large");
    chunks.push(chunk);
  }
  return Buffer.concat(chunks);
}

async function proxy(req, res, pathname) {
  const method = API_ROUTES[pathname];
  if (req.method !== method) throw new HttpError(405, `Use ${method}`);
  const auth = req.headers.authorization;
  if (!auth) throw new HttpError(401, "Missing API key: set your TypeSafe key in the Jev pane");

  const headers = { Accept: "application/json", Authorization: auth };
  if (method === "POST") headers["Content-Type"] = "application/json";

  const upstream = await fetch(UPSTREAM + pathname, {
    method,
    headers,
    body: method === "POST" ? await readBody(req) : undefined,
    signal: AbortSignal.timeout(UPSTREAM_TIMEOUT_MS),
  });
  const out = { "Cache-Control": "no-store" };
  for (const name of PASS_HEADERS) {
    const value = upstream.headers.get(name);
    if (value) out[name] = value;
  }
  res.writeHead(upstream.status, out);
  res.end(Buffer.from(await upstream.arrayBuffer()));
}

/** Map a URL path to a file under ROOT, or null if it's malformed or escapes ROOT. */
function resolveFile(pathname) {
  let decoded;
  try {
    decoded = decodeURIComponent(pathname);
  } catch (_) {
    return null;
  }
  if (decoded === "/") decoded = "/taskpane.html";
  const file = path.resolve(ROOT, "." + decoded);
  return file.startsWith(ROOT + path.sep) ? file : null;
}

function serveStatic(res, pathname) {
  const file = resolveFile(pathname);
  if (!file) {
    res.writeHead(400, { "Content-Type": "text/plain" }).end("Bad request");
    return;
  }
  fs.readFile(file, (err, data) => {
    if (err) {
      res.writeHead(404, { "Content-Type": "text/plain" }).end("Not found");
      return;
    }
    res.writeHead(200, {
      "Content-Type": TYPES[path.extname(file).toLowerCase()] || "application/octet-stream",
      "Cache-Control": "no-store",
      // Excel on the web fetches functions.json from another origin.
      "Access-Control-Allow-Origin": "*",
    });
    res.end(data);
  });
}

function handle(req, res) {
  let pathname;
  try {
    ({ pathname } = new URL(req.url, ORIGIN));
  } catch (_) {
    res.writeHead(400, { "Content-Type": "text/plain" }).end("Bad request");
    return;
  }
  if (!API_ROUTES[pathname]) {
    serveStatic(res, pathname);
    return;
  }
  proxy(req, res, pathname).catch((err) => {
    if (res.headersSent) {
      res.destroy();
      return;
    }
    if (err instanceof HttpError) sendJson(res, err.status, { error: { message: err.message } });
    else sendJson(res, 502, { error: { message: `Dev proxy could not reach ${UPSTREAM}: ${err.message}` } });
  });
}

async function main() {
  const { getHttpsServerOptions } = require("office-addin-dev-certs");
  const options = await getHttpsServerOptions();
  https.createServer(options, handle).listen(PORT, () => {
    console.log(`Jev for Excel dev server: ${ORIGIN}/taskpane.html`);
    console.log(`Proxying ${Object.keys(API_ROUTES).join(" and ")} to ${UPSTREAM}`);
  });
}

if (require.main === module) {
  main().catch((err) => {
    console.error(err.message || err);
    console.error("Run `npm run certs` first to create and trust the localhost certificate.");
    process.exit(1);
  });
}

module.exports = { handle, resolveFile, ROOT };
