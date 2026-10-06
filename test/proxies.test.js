"use strict";
/* Tests for the dev-server proxy (scripts/serve.js) and the Cloudflare Worker (worker/index.mjs). */
const test = require("node:test");
const assert = require("node:assert/strict");
const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");
const { handle, resolveFile, ROOT } = require("../scripts/serve.js");

/** Replace global fetch for one test, recording the upstream calls. */
function stubFetch(t, respond = () => new Response('{"ok":true}', { headers: { "Content-Type": "application/json" } })) {
  const calls = [];
  const original = globalThis.fetch;
  globalThis.fetch = async (url, init) => {
    calls.push({ url, init });
    return respond(url, init);
  };
  t.after(() => (globalThis.fetch = original));
  return calls;
}

const USER_KEY = { Authorization: "Bearer sk-user" };

// ---------------------------------------------------------------------------
// Dev server
// ---------------------------------------------------------------------------

test("dev server: static paths stay inside src/", () => {
  assert.equal(resolveFile("/"), path.join(ROOT, "taskpane.html"));
  assert.equal(resolveFile("/assets/icon-16.png"), path.join(ROOT, "assets", "icon-16.png"));
  assert.equal(resolveFile("/../package.json"), null);
  assert.equal(resolveFile("/%2e%2e/package.json"), null);
  assert.equal(resolveFile("/../src-other/secret.txt"), null, "sibling folder with a matching prefix");
  assert.equal(resolveFile("/%E0"), null, "malformed escape");
});

async function withDevServer(t) {
  const server = http.createServer(handle);
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  t.after(() => server.close());
  const base = `http://127.0.0.1:${server.address().port}`;
  // Capture the real fetch before any stub, so requests reach the server.
  const realFetch = globalThis.fetch;
  return (p, init) => realFetch(base + p, init);
}

test("dev server: proxy passes the caller's key through", async (t) => {
  const request = await withDevServer(t);
  const calls = stubFetch(t);
  const res = await request("/v1/systemone", {
    method: "POST",
    headers: { "Content-Type": "application/json", ...USER_KEY },
    body: '{"model":"jev-latest"}',
  });
  assert.equal(res.status, 200);
  assert.equal(calls.length, 1);
  assert.equal(calls[0].url, "https://api.typesafe.ai/v1/systemone");
  assert.equal(calls[0].init.headers.Authorization, "Bearer sk-user");
  assert.equal(Buffer.from(calls[0].init.body).toString(), '{"model":"jev-latest"}');
});

test("dev server: proxy refuses requests without a key or with the wrong method", async (t) => {
  const request = await withDevServer(t);
  const calls = stubFetch(t);
  assert.equal((await request("/v1/systemone", { method: "POST", body: "{}" })).status, 401);
  assert.equal((await request("/v1/systemone", { headers: USER_KEY })).status, 405);
  assert.equal(calls.length, 0, "nothing reached TypeSafe");
});

// ---------------------------------------------------------------------------
// Cloudflare Worker
// ---------------------------------------------------------------------------

const loadWorker = async () => (await import("../worker/index.mjs")).default;
const workerRequest = (p, init = {}) => new Request("https://jev.example.workers.dev" + p, init);

/** Stands in for Cloudflare's static-asset binding, serving the dev manifest and functions.json. */
const ENV = {
  ASSETS: {
    fetch: async (request) => {
      const { pathname } = new URL(request.url);
      if (pathname === "/manifest.xml") return new Response(fs.readFileSync(path.join(__dirname, "..", "manifest.xml")));
      if (pathname === "/functions.json") return new Response(fs.readFileSync(path.join(ROOT, "functions.json")));
      return new Response("Not found", { status: 404 });
    },
  },
};

test("worker: passes the caller's key through", async (t) => {
  const worker = await loadWorker();
  const calls = stubFetch(t);
  const res = await worker.fetch(
    workerRequest("/v1/systemone", {
      method: "POST",
      headers: { ...USER_KEY, "Content-Type": "application/json" },
      body: "{}",
    }),
    ENV
  );
  assert.equal(res.status, 200);
  assert.equal(calls[0].url, "https://api.typesafe.ai/v1/systemone");
  assert.equal(calls[0].init.headers.Authorization, "Bearer sk-user");
});

test("worker: rejects missing keys, wrong methods and oversize bodies", async (t) => {
  const worker = await loadWorker();
  const calls = stubFetch(t);
  const status = async (req) => (await worker.fetch(req, ENV)).status;

  assert.equal(await status(workerRequest("/v1/models")), 401);
  assert.equal(await status(workerRequest("/v1/models", { method: "POST", headers: USER_KEY, body: "{}" })), 405);
  const big = "x".repeat(1_000_001);
  assert.equal(await status(workerRequest("/v1/systemone", { method: "POST", headers: USER_KEY, body: big })), 413);
  assert.equal(calls.length, 0, "nothing reached TypeSafe");
});

test("worker: reports upstream failures as 502", async (t) => {
  const worker = await loadWorker();
  stubFetch(t, () => {
    throw new TypeError("network down");
  });
  const res = await worker.fetch(workerRequest("/v1/models", { headers: USER_KEY }), ENV);
  assert.equal(res.status, 502);
});

test("worker: serves the manifest with its own origin and leaves other paths to the assets", async () => {
  const worker = await loadWorker();
  const res = await worker.fetch(workerRequest("/manifest.xml"), ENV);
  const xml = await res.text();
  assert.match(res.headers.get("Content-Type"), /xml/);
  assert.doesNotMatch(xml, /localhost:3000/);
  assert.ok(xml.includes('"https://jev.example.workers.dev/functions.json"'));
  assert.equal((await worker.fetch(workerRequest("/nope.js"), ENV)).status, 404);
});

test("worker: serves functions.json with help links on its own origin, readable cross-origin", async () => {
  const worker = await loadWorker();
  const res = await worker.fetch(workerRequest("/functions.json"), ENV);
  const { functions } = await res.json();
  assert.match(res.headers.get("Content-Type"), /json/);
  assert.equal(res.headers.get("Access-Control-Allow-Origin"), "*", "Excel on the web reads it from another origin");
  for (const fn of functions) {
    assert.ok(fn.helpUrl.startsWith("https://jev.example.workers.dev/cheatsheet.html#"), `${fn.id}: ${fn.helpUrl}`);
  }
});
