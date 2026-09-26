"use strict";
/* Unit tests for src/jev-core.js — run with `npm test` (Node 20+). No network needed. */
const test = require("node:test");
const assert = require("node:assert/strict");
const Jev = require("../src/jev-core.js");

// ---------------------------------------------------------------------------
// A fake Jev API: answers every question deterministically and records calls.
// ---------------------------------------------------------------------------
function fakeApi({ status = 200, failTimes = 0, failStatus = 429, headers = {}, rejectIf } = {}) {
  const calls = [];
  let failures = 0;
  const fetch = async (url, init) => {
    const body = init.body ? JSON.parse(init.body) : undefined;
    calls.push({ url, init, body });
    const mkRes = (code, obj, extra = {}) => ({
      ok: code >= 200 && code < 300,
      status: code,
      headers: new Map(Object.entries({ ...extra })),
      text: async () => (typeof obj === "string" ? obj : JSON.stringify(obj)),
    });
    if (failures < failTimes) {
      failures++;
      return mkRes(failStatus, { error: { message: "slow down" } }, headers);
    }
    if (status !== 200) return mkRes(status, { detail: "nope" });
    if (url.endsWith("/v1/models")) {
      return mkRes(200, { models: [{ name: "jev-1.13.0", description: "System One", release_date: "2026-09-18" }] });
    }
    if (rejectIf && rejectIf(body)) return mkRes(422, { detail: [{ loc: ["body", "questions"], msg: "bad question" }] });
    const answers = {};
    for (const [id, q] of Object.entries(body.questions)) {
      if (q.type === "noul") answers[id] = { type: "noul", noul: 0.82 };
      if (q.type === "choice") {
        const labels = Object.keys(q.criteria);
        const probabilities = {};
        labels.forEach((l, i) => (probabilities[l] = i === 1 ? 0.7 : 0.3 / (labels.length - 1)));
        answers[id] = { type: "choice", choice: labels[1], confidence: 0.55, probabilities };
      }
      if (q.type === "score") {
        const legend = {};
        const probabilities = {};
        q.criteria.forEach((c, i) => {
          legend[String(i)] = c;
          probabilities[String(i)] = i === 1 ? 0.57 : i === 2 ? 0.43 : 0;
        });
        answers[id] = { type: "score", score: 1.43, confidence: 0.35, legend, probabilities };
      }
    }
    return mkRes(200, { model: "jev-1.13.0", answers, usage: { input_tokens: 100, output_tokens: 3 } });
  };
  // Map has .get() like Headers; make missing keys return null like Headers does
  const wrap = async (url, init) => {
    const res = await fetch(url, init);
    const m = res.headers;
    res.headers = { get: (k) => (m.has(k) ? m.get(k) : null) };
    return res;
  };
  return { fetch: wrap, calls };
}

const client = (api, extra = {}) =>
  new Jev.JevClient({ apiKey: "sk-test", fetch: api.fetch, batchWindowMs: 5, ...extra });

// ---------------------------------------------------------------------------
test("parseMode: types, aliases, outputs and thresholds", () => {
  assert.deepEqual(Jev.parseMode("choice"), { type: "choice", output: "label", threshold: 0.5 });
  assert.deepEqual(Jev.parseMode(" CHOICE:Probs "), { type: "choice", output: "probs", threshold: 0.5 });
  assert.equal(Jev.parseMode("classify").type, "choice");
  assert.equal(Jev.parseMode("yesno").type, "noul");
  assert.equal(Jev.parseMode("rate:label").output, "label");
  assert.deepEqual(Jev.parseMode("noul:bool:0.8"), { type: "noul", output: "bool", threshold: 0.8 });
  assert.throws(() => Jev.parseMode("banana"), /Unknown mode/);
  assert.throws(() => Jev.parseMode("choice:level"), /Unknown output/);
  assert.throws(() => Jev.parseMode("choice:label:0.5"), /threshold/i);
  assert.throws(() => Jev.parseMode("noul:bool:1.5"), /between 0 and 1/);
  assert.throws(() => Jev.parseMode(""), /required/);
});

test("normalizeState: text, numbers, JSON and ranges", () => {
  assert.equal(Jev.normalizeState([["Hello"]]), "Hello");
  assert.equal(Jev.normalizeState(42), "42");
  assert.equal(Jev.normalizeState([[""]]), null);
  assert.equal(Jev.normalizeState(null), null);
  assert.deepEqual(Jev.normalizeState('{"a":1}'), { a: 1 });
  assert.equal(Jev.normalizeState("{not json}"), "{not json}");
  assert.deepEqual(Jev.normalizeState([["a", ""], ["b", "c"]]), ["a", "b", "c"]);
});

test("buildStateJson pairs headers and values", () => {
  const json = Jev.buildStateJson([["name", "amount", ""]], [["Acme", 120, "ignored"]]);
  assert.deepEqual(JSON.parse(json), { name: "Acme", amount: 120 });
  assert.throws(() => Jev.buildStateJson([["a", "b"]], [["x"]]), /same number/);
});

test("buildQuestion: choice with labels, pipes, pairs and ?question", () => {
  const { question, meta } = Jev.buildQuestion("choice", [
    [["?Which team should handle this?"]],
    [["billing | Payments and refunds"]],
    [["technical"]],
    [
      ["sales", "Pricing questions"],
      ["other", ""],
    ],
  ]);
  assert.deepEqual(question, {
    type: "choice",
    instructions: "Which team should handle this?",
    criteria: { billing: "Payments and refunds", technical: null, sales: "Pricing questions", other: null },
  });
  assert.deepEqual(meta.labels, ["billing", "technical", "sales", "other"]);
});

test("buildQuestion: a single row of two cells is two labels, not a pair", () => {
  const { question } = Jev.buildQuestion("choice", [[["yes", "no"]]]);
  assert.deepEqual(question.criteria, { yes: null, no: null });
  assert.equal("instructions" in question, false);
});

test("buildQuestion: choice validation", () => {
  assert.throws(() => Jev.buildQuestion("choice", ["only"]), /at least 2/);
  assert.throws(() => Jev.buildQuestion("choice", ["a", "a"]), /Duplicate/);
  const many = Array.from({ length: 256 }, (_, i) => "o" + i);
  assert.throws(() => Jev.buildQuestion("choice", [many]), /at most 255/);
  assert.throws(() => Jev.buildQuestion("choice", ["?q1", "?q2", "a", "b"]), /Only one/);
});

test("buildQuestion: score levels", () => {
  const { question } = Jev.buildQuestion("score", ["?How angry?", [["Calm"], ["Annoyed"], ["Furious"]]]);
  assert.deepEqual(question, { type: "score", instructions: "How angry?", criteria: ["Calm", "Annoyed", "Furious"] });
  assert.throws(() => Jev.buildQuestion("score", ["one"]), /at least 2/);
  assert.throws(() => Jev.buildQuestion("score", [Array.from({ length: 11 }, String)]), /at most 10/);
});

test("buildQuestion: noul question and criteria", () => {
  assert.deepEqual(Jev.buildQuestion("noul", ["Asks for a refund"]).question, {
    type: "noul",
    instructions: "Asks for a refund",
  });
  assert.deepEqual(Jev.buildQuestion("noul", ["Urgent?", "Mentions a deadline", "No time pressure"]).question, {
    type: "noul",
    instructions: "Urgent?",
    criteria: { true: "Mentions a deadline", false: "No time pressure" },
  });
  // "?" prefix works too, and then remaining items are criteria
  assert.deepEqual(Jev.buildQuestion("noul", ["Deadline stated", "?Is it urgent?"]).question, {
    type: "noul",
    instructions: "Is it urgent?",
    criteria: { true: "Deadline stated" },
  });
  // NOUL(text, question, , whenFalse) passes null for the gap
  assert.deepEqual(Jev.buildQuestion("noul", ["Q", null, null]).question, { type: "noul", instructions: "Q" });
  assert.throws(() => Jev.buildQuestion("noul", []), /needs a yes\/no question/);
  assert.throws(() => Jev.buildQuestion("noul", ["q", "a", "b", "c"]), /at most two/);
});

test("formatAnswer: every output", () => {
  const choice = { type: "choice", choice: "billing", confidence: 0.6, probabilities: { billing: 0.84, technical: 0.15, other: 0.01 } };
  const f = (a, mode, meta) => Jev.formatAnswer(a, Jev.parseMode(mode), meta);
  assert.deepEqual(f(choice, "choice"), [["billing"]]);
  assert.deepEqual(f(choice, "choice:confidence"), [[0.6]]);
  assert.deepEqual(f(choice, "choice:prob"), [[0.84]]);
  assert.deepEqual(f(choice, "choice:probs"), [["billing", 0.84], ["technical", 0.15], ["other", 0.01]]);
  assert.deepEqual(JSON.parse(f(choice, "choice:json")[0][0]), choice);

  const score = {
    type: "score", score: 1.43, confidence: 0.35,
    legend: { 0: "Cosmetic", 1: "Degraded", 2: "Blocking" },
    probabilities: { 2: 0.43, 0: 0, 1: 0.57 },
  };
  const meta = { levels: ["Cosmetic", "Degraded", "Blocking"] };
  assert.deepEqual(f(score, "score", meta), [[1.43]]);
  assert.deepEqual(f(score, "score:level", meta), [[1]]);
  assert.deepEqual(f(score, "score:label", meta), [["Degraded"]]);
  assert.deepEqual(f(score, "score:probs", meta), [[0, 0], [1, 0.57], [2, 0.43]]);

  const noul = { type: "noul", noul: 0.72 };
  assert.deepEqual(f(noul, "noul"), [[0.72]]);
  assert.deepEqual(f(noul, "noul:bool"), [[true]]);
  assert.deepEqual(f(noul, "noul:bool:0.8"), [[false]]);
  assert.deepEqual(f(noul, "noul:yesno"), [["Yes"]]);
  assert.throws(() => f({ type: "noul" }, "noul"), /missing noul/);
});

test("evaluate end-to-end builds the right request", async () => {
  const api = fakeApi();
  const c = client(api, { model: "jev-1.13.0" });
  const out = await Jev.evaluate(c, "choice", [["My card was charged twice"]], ["billing", "technical", "other"]);
  assert.deepEqual(out, [["technical"]]);
  assert.equal(api.calls.length, 1);
  const { url, init, body } = api.calls[0];
  assert.equal(url, "https://api.typesafe.ai/v1/systemone");
  assert.equal(init.method, "POST");
  assert.equal(init.headers.Authorization, "Bearer sk-test");
  assert.equal(init.headers["Content-Type"], "application/json");
  assert.deepEqual(body, {
    model: "jev-1.13.0",
    state: "My card was charged twice",
    questions: { q0: { type: "choice", criteria: { billing: null, technical: null, other: null } } },
  });
});

test("blank text returns blank without calling the API", async () => {
  const api = fakeApi();
  const out = await Jev.evaluate(client(api), "noul", [[""]], ["Is it urgent?"]);
  assert.deepEqual(out, [[""]]);
  assert.equal(api.calls.length, 0);
});

test("questions about the same text are batched into one request", async () => {
  const api = fakeApi();
  const c = client(api);
  const text = [["Server is down and customers cannot log in!"]];
  const results = await Promise.all([
    Jev.evaluate(c, "noul", text, ["Is this urgent?"]),
    Jev.evaluate(c, "choice", text, ["billing", "technical"]),
    Jev.evaluate(c, "score:level", text, ["low", "medium", "high"]),
    Jev.evaluate(c, "noul", [["Different text"]], ["Is this urgent?"]),
  ]);
  assert.deepEqual(results, [[[0.82]], [["technical"]], [[1]], [[0.82]]]);
  assert.equal(api.calls.length, 2, "one request per distinct text");
  const sizes = api.calls.map((call) => Object.keys(call.body.questions).length).sort();
  assert.deepEqual(sizes, [1, 3]);
  assert.equal(c.stats.inputTokens, 200);
});

test("cache: repeated and output-variant calls reuse one answer", async () => {
  const api = fakeApi();
  const c = client(api);
  const args = ["billing", "technical", "other"];
  await Jev.evaluate(c, "choice", "hello", args);
  await Jev.evaluate(c, "choice:probs", "hello", args); // same question, different output
  await Jev.evaluate(c, "choice:confidence", "hello", args);
  assert.equal(api.calls.length, 1);
  assert.equal(c.stats.cacheHits, 2);
  c.clearCache();
  await Jev.evaluate(c, "choice", "hello", args);
  assert.equal(api.calls.length, 2);
});

test("cache disabled: every call hits the API", async () => {
  const api = fakeApi();
  const c = client(api, { cache: false });
  await Jev.evaluate(c, "noul", "x", ["q"]);
  await Jev.evaluate(c, "noul", "x", ["q"]);
  assert.equal(api.calls.length, 2);
});

test("retries 429 honouring retry-after-ms, then succeeds", async () => {
  const api = fakeApi({ failTimes: 2, failStatus: 429, headers: { "retry-after-ms": "1" } });
  const c = client(api);
  const out = await Jev.evaluate(c, "noul", "x", ["q"]);
  assert.deepEqual(out, [[0.82]]);
  assert.equal(api.calls.length, 3);
});

test("gives up after maxRetries with a readable error", async () => {
  const api = fakeApi({ failTimes: 99, failStatus: 529, headers: { "retry-after-ms": "1" } });
  const c = client(api, { maxRetries: 1 });
  await assert.rejects(Jev.evaluate(c, "noul", "x", ["q"]), (e) => e.status === 529 && /overloaded/.test(e.message));
  assert.equal(api.calls.length, 2);
});

test("401 maps to an API-key message and is not cached", async () => {
  const api = fakeApi({ status: 401 });
  const c = client(api);
  await assert.rejects(Jev.evaluate(c, "noul", "x", ["q"]), /API key/);
  assert.equal(c.cache.size, 0);
});

test("missing API key fails fast, whatever the base URL", async () => {
  const api = fakeApi();
  for (const baseUrl of ["https://api.typesafe.ai", "https://jev-proxy.example.workers.dev"]) {
    const c = new Jev.JevClient({ fetch: api.fetch, baseUrl });
    await assert.rejects(Jev.evaluate(c, "noul", "x", ["q"]), /No Jev API key/);
    await assert.rejects(c.listModels(), /No Jev API key/);
  }
  assert.equal(api.calls.length, 0);
});

test("proxy base URL gets the user's key and trims trailing slash", async () => {
  const api = fakeApi();
  const c = client(api, { baseUrl: "https://jev-proxy.example.workers.dev/", batchWindowMs: 1 });
  await Jev.evaluate(c, "noul", "x", ["q"]);
  assert.equal(api.calls[0].url, "https://jev-proxy.example.workers.dev/v1/systemone");
  assert.equal(api.calls[0].init.headers.Authorization, "Bearer sk-test");
});

test("a 422 in a batch is retried per question so only the bad cell errors", async () => {
  const api = fakeApi({ rejectIf: (body) => Object.values(body.questions).some((q) => q.instructions === "BAD") });
  const c = client(api);
  const [good, bad] = await Promise.allSettled([
    Jev.evaluate(c, "noul", "same text", ["fine question"]),
    Jev.evaluate(c, "noul", "same text", ["BAD"]),
  ]);
  assert.equal(good.status, "fulfilled");
  assert.deepEqual(good.value, [[0.82]]);
  assert.equal(bad.status, "rejected");
  assert.match(bad.reason.message, /rejected the question.*bad question/);
});

test("listModels reads { models: [...] }", async () => {
  const api = fakeApi();
  const models = await client(api).listModels();
  assert.equal(models[0].name, "jev-1.13.0");
  assert.equal(api.calls[0].init.method, "GET");
  assert.equal(api.calls[0].init.body, undefined);
});

test("network failure produces a CORS/proxy hint", async () => {
  const c = new Jev.JevClient({
    apiKey: "k",
    batchWindowMs: 1,
    maxRetries: 0,
    fetch: async () => {
      throw new TypeError("Failed to fetch");
    },
  });
  await assert.rejects(Jev.evaluate(c, "noul", "x", ["q"]), /network or CORS/);
});

test("maxBatchSize splits large batches", async () => {
  const api = fakeApi();
  const c = client(api, { maxBatchSize: 2 });
  await Promise.all(["a", "b", "c", "d", "e"].map((q) => Jev.evaluate(c, "noul", "same", [q])));
  assert.equal(api.calls.length, 3);
});

test("concurrency is limited", async () => {
  let inFlight = 0;
  let peak = 0;
  const api = fakeApi();
  const slowFetch = async (url, init) => {
    inFlight++;
    peak = Math.max(peak, inFlight);
    await new Promise((r) => setTimeout(r, 10));
    try {
      return await api.fetch(url, init);
    } finally {
      inFlight--;
    }
  };
  const c = new Jev.JevClient({ apiKey: "k", fetch: slowFetch, batchWindowMs: 1, maxConcurrency: 3 });
  await Promise.all(Array.from({ length: 12 }, (_, i) => Jev.evaluate(c, "noul", "text " + i, ["q"])));
  assert.equal(api.calls.length, 12);
  assert.ok(peak <= 3, `peak concurrency ${peak}`);
});

test("parseMode: too many parts is reported as such", () => {
  assert.throws(() => Jev.parseMode("choice:label:x:y"), /too many parts/);
  assert.throws(() => Jev.parseMode("choice:label:0.5"), /only valid with noul/);
});

test("buildQuestion: a ?question inside a label/description range is the question", () => {
  const { question } = Jev.buildQuestion("choice", [
    [
      ["?Which team?", ""],
      ["billing", "Payments"],
      ["other", ""],
    ],
  ]);
  assert.equal(question.instructions, "Which team?");
  assert.deepEqual(question.criteria, { billing: "Payments", other: null });
});

test("network failure against api.typesafe.ai explains that a proxy is needed", async () => {
  const failing = async () => {
    throw new TypeError("Failed to fetch");
  };
  const direct = new Jev.JevClient({ apiKey: "k", batchWindowMs: 1, maxRetries: 0, fetch: failing });
  await assert.rejects(Jev.evaluate(direct, "noul", "x", ["q"]), /blocks calls from browser-based add-ins/);
  const proxied = new Jev.JevClient({
    apiKey: "k",
    baseUrl: "https://proxy.example/",
    batchWindowMs: 1,
    maxRetries: 0,
    fetch: failing,
  });
  await assert.rejects(Jev.evaluate(proxied, "noul", "x", ["q"]), /Could not reach https:\/\/proxy\.example \(/);
});
