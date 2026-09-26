/*!
 * Jev for Excel — core logic.
 *
 * Shared by the Office add-in (loaded as a browser global, `JevCore`) and the
 * Node test-suite (loaded via `require`). No DOM or Office.js dependencies.
 *
 * Talks to TypeSafe AI's System One endpoint:
 *   POST {baseUrl}/v1/systemone   { model, state, questions }
 *   GET  {baseUrl}/v1/models
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.JevCore = api;
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const VERSION = "1.0.0";

  const DEFAULTS = Object.freeze({
    apiKey: "",
    baseUrl: "https://api.typesafe.ai",
    model: "jev-latest",
    cache: true,
    batchWindowMs: 40, // gather calls for this long before sending
    maxBatchSize: 16, // max questions per request (same state + model)
    maxConcurrency: 6, // max requests in flight
    timeoutMs: 15000, // per attempt
    maxRetries: 3, // on 408 / 429 / 5xx / network errors
    maxCacheEntries: 20000,
    fetch: null, // injectable for tests / custom transports
  });

  const LIMITS = Object.freeze({ choiceMin: 2, choiceMax: 255, scoreMin: 2, scoreMax: 10 });

  // ---------------------------------------------------------------------------
  // Errors
  // ---------------------------------------------------------------------------

  /** An error with a user-facing message. `kind` maps to an Excel error value. */
  class JevError extends Error {
    constructor(message, kind = "value", status) {
      super(message);
      this.name = "JevError";
      this.kind = kind; // "value" -> #VALUE!, "na" -> #N/A
      if (status !== undefined) this.status = status;
    }
  }

  // ---------------------------------------------------------------------------
  // Mode parsing:  "type[:output[:threshold]]"
  // ---------------------------------------------------------------------------

  const TYPE_ALIASES = {
    choice: "choice", choose: "choice", classify: "choice", category: "choice", pick: "choice", select: "choice",
    score: "score", rate: "score", rating: "score", scale: "score", rubric: "score",
    noul: "noul", yesno: "noul", yn: "noul", bool: "noul", boolean: "noul", binary: "noul",
  };

  const OUTPUT_ALIASES = {
    choice: {
      "": "label", label: "label", choice: "label",
      confidence: "confidence", conf: "confidence",
      prob: "prob", probability: "prob",
      probs: "probs", probabilities: "probs",
      json: "json",
    },
    score: {
      "": "score", score: "score",
      level: "level", round: "level",
      label: "label", legend: "label",
      confidence: "confidence", conf: "confidence",
      probs: "probs", probabilities: "probs",
      json: "json",
    },
    noul: {
      "": "prob", prob: "prob", probability: "prob", noul: "prob",
      bool: "bool", boolean: "bool",
      yesno: "yesno",
      json: "json",
    },
  };

  function parseMode(mode) {
    if (isEmpty(mode)) {
      throw new JevError('Mode is required: "choice", "score" or "noul".');
    }
    const parts = String(mode).trim().toLowerCase().split(":").map((s) => s.trim());
    if (parts.length > 3) throw new JevError(`Mode "${mode}" has too many parts.`);
    const type = TYPE_ALIASES[parts[0]];
    if (!type) {
      throw new JevError(`Unknown mode "${parts[0]}". Use "choice", "score" or "noul".`);
    }
    const outRaw = parts[1] || "";
    const output = OUTPUT_ALIASES[type][outRaw];
    if (!output) {
      const valid = Object.keys(OUTPUT_ALIASES[type]).filter(Boolean).join(", ");
      throw new JevError(`Unknown output "${outRaw}" for ${type}. Valid: ${valid}.`);
    }
    let threshold = 0.5;
    if (parts[2] !== undefined && parts[2] !== "") {
      if (type !== "noul" || (output !== "bool" && output !== "yesno")) {
        throw new JevError("A threshold is only valid with noul:bool or noul:yesno.");
      }
      threshold = Number(parts[2]);
      if (!Number.isFinite(threshold) || threshold <= 0 || threshold >= 1) {
        throw new JevError("Threshold must be a number between 0 and 1, e.g. noul:bool:0.7.");
      }
    }
    return { type, output, threshold };
  }

  // ---------------------------------------------------------------------------
  // Value helpers
  // ---------------------------------------------------------------------------

  function isEmpty(v) {
    return v === null || v === undefined || (typeof v === "string" && v.trim() === "");
  }

  function toText(v) {
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    if (v === null || v === undefined) return "";
    return String(v);
  }

  /** Flatten a scalar, a 1-D array or a 2-D matrix (row-major). */
  function flatten(v) {
    if (!Array.isArray(v)) return [v];
    const out = [];
    for (const item of v) {
      if (Array.isArray(item)) out.push(...item);
      else out.push(item);
    }
    return out;
  }

  function isMatrix(v) {
    return Array.isArray(v) && v.length > 0 && Array.isArray(v[0]);
  }

  function maybeJson(s) {
    const t = s.trim();
    const looksJson = (t.startsWith("{") && t.endsWith("}")) || (t.startsWith("[") && t.endsWith("]"));
    if (looksJson) {
      try {
        return JSON.parse(t);
      } catch (_) {
        /* not JSON — treat as plain text */
      }
    }
    return s;
  }

  // ---------------------------------------------------------------------------
  // State (the text Jev evaluates)
  // ---------------------------------------------------------------------------

  /**
   * Single cell/string -> string (or parsed JSON object/array if it is valid JSON).
   * Multi-cell range  -> array of the non-empty cell values as text.
   * Empty             -> null (caller should short-circuit and not call the API).
   */
  function normalizeState(text) {
    const cells = flatten(text).filter((v) => !isEmpty(v)).map(toText);
    if (cells.length === 0) return null;
    if (cells.length === 1) return maybeJson(cells[0]);
    return cells;
  }

  /** JEV.STATE(headers, values) -> JSON object string, for named-field state. */
  function buildStateJson(headers, values) {
    const h = flatten(headers);
    const v = flatten(values);
    if (h.length !== v.length) {
      throw new JevError(`STATE needs the same number of headers (${h.length}) and values (${v.length}).`);
    }
    const obj = {};
    h.forEach((key, i) => {
      if (isEmpty(key)) return;
      const val = v[i];
      obj[toText(key).trim()] = val === undefined || val === "" ? null : val;
    });
    if (Object.keys(obj).length === 0) throw new JevError("STATE needs at least one non-empty header.");
    return JSON.stringify(obj);
  }

  // ---------------------------------------------------------------------------
  // Questions
  // ---------------------------------------------------------------------------

  /** "?Which team?" -> true. A lone "?" is an ordinary option. */
  function isQuestion(s) {
    return s.length > 1 && s.startsWith("?");
  }

  /**
   * Turn the variadic option arguments into a Jev question.
   *  - An item starting with "?" is the question (instructions) itself.
   *  - choice: items are labels; "label | description" adds a description;
   *            a range with 2+ rows and exactly 2 columns is read as label/description pairs.
   *  - score:  items are the rubric levels in order (level 0 first), 2–10 of them.
   *  - noul:   the first non-"?" item is the question if no "?" item was given;
   *            up to two further items describe the TRUE and FALSE outcomes.
   */
  function buildQuestion(type, args) {
    const questions = [];
    const options = []; // { text, description? }

    for (const arg of args || []) {
      if (arg === null || arg === undefined) continue;
      if (type === "choice" && isMatrix(arg) && arg.length >= 2 && arg[0].length === 2) {
        for (const row of arg) {
          if (isEmpty(row[0])) continue;
          const first = toText(row[0]).trim();
          if (isQuestion(first)) {
            questions.push(first.slice(1).trim());
            continue;
          }
          options.push({
            text: first,
            description: isEmpty(row[1]) ? null : toText(row[1]).trim(),
          });
        }
        continue;
      }
      for (const v of flatten(arg)) {
        if (isEmpty(v)) continue;
        const s = toText(v).trim();
        if (isQuestion(s)) questions.push(s.slice(1).trim());
        else options.push({ text: s });
      }
    }

    if (questions.length > 1) throw new JevError('Only one "?question" argument is allowed.');
    let instructions = questions.length ? questions[0] : null;

    if (type === "choice") {
      const criteria = {};
      const labels = [];
      for (const o of options) {
        let label = o.text;
        let description = o.description;
        if (description === undefined) {
          const i = label.indexOf("|");
          if (i >= 0) {
            description = label.slice(i + 1).trim() || null;
            label = label.slice(0, i).trim();
          } else {
            description = null;
          }
        }
        if (!label) throw new JevError("Choice labels cannot be empty.");
        if (Object.prototype.hasOwnProperty.call(criteria, label)) {
          throw new JevError(`Duplicate choice label "${label}".`);
        }
        criteria[label] = description;
        labels.push(label);
      }
      if (labels.length < LIMITS.choiceMin) throw new JevError("choice needs at least 2 options.");
      if (labels.length > LIMITS.choiceMax) throw new JevError("choice supports at most 255 options.");
      const question = { type: "choice" };
      if (instructions !== null) question.instructions = instructions;
      question.criteria = criteria;
      return { question, meta: { labels } };
    }

    if (type === "score") {
      const levels = options.map((o) => o.text);
      if (levels.length < LIMITS.scoreMin) throw new JevError("score needs at least 2 rubric levels.");
      if (levels.length > LIMITS.scoreMax) throw new JevError("score supports at most 10 rubric levels.");
      const question = { type: "score" };
      if (instructions !== null) question.instructions = instructions;
      question.criteria = levels;
      return { question, meta: { levels } };
    }

    if (type === "noul") {
      const rest = options.map((o) => o.text);
      if (instructions === null) {
        if (!rest.length) throw new JevError("noul needs a yes/no question, e.g. \"Is the customer asking for a refund?\"");
        instructions = rest.shift();
      }
      if (rest.length > 2) {
        throw new JevError("noul takes a question plus at most two descriptions (TRUE, then FALSE).");
      }
      const question = { type: "noul", instructions };
      if (rest.length) {
        question.criteria = { true: rest[0] };
        if (rest.length > 1) question.criteria.false = rest[1];
      }
      return { question, meta: {} };
    }

    throw new JevError(`Unsupported type "${type}".`);
  }

  // ---------------------------------------------------------------------------
  // Formatting answers for Excel (always a 2-D array)
  // ---------------------------------------------------------------------------

  function requireNumber(v, what) {
    const n = typeof v === "number" ? v : Number(v);
    if (!Number.isFinite(n)) throw new JevError(`Jev response is missing ${what}.`, "na");
    return n;
  }

  function sortedProbs(probs) {
    return Object.entries(probs || {})
      .map(([k, p]) => [k, Number(p)])
      .sort((a, b) => b[1] - a[1]);
  }

  function formatAnswer(answer, mode, meta = {}) {
    if (!answer || typeof answer !== "object") throw new JevError("Empty answer from Jev.", "na");
    if (mode.output === "json") return [[JSON.stringify(answer)]];

    if (mode.type === "choice") {
      switch (mode.output) {
        case "label":
          if (typeof answer.choice !== "string") throw new JevError("Jev response is missing choice.", "na");
          return [[answer.choice]];
        case "confidence":
          return [[requireNumber(answer.confidence, "confidence")]];
        case "prob":
          return [[requireNumber(answer.probabilities && answer.probabilities[answer.choice], "probabilities")]];
        case "probs": {
          const rows = sortedProbs(answer.probabilities);
          if (!rows.length) throw new JevError("Jev response is missing probabilities.", "na");
          return rows;
        }
      }
    }

    if (mode.type === "score") {
      const score = requireNumber(answer.score, "score");
      const maxLevel = meta.levels ? meta.levels.length - 1 : Infinity;
      const level = Math.min(Math.max(Math.round(score), 0), maxLevel);
      switch (mode.output) {
        case "score":
          return [[score]];
        case "level":
          return [[level]];
        case "label": {
          const legend = answer.legend || {};
          const text = legend[level] ?? (meta.levels ? meta.levels[level] : undefined);
          if (text === undefined || text === null) return [[level]];
          return [[typeof text === "string" ? text : JSON.stringify(text)]];
        }
        case "confidence":
          return [[requireNumber(answer.confidence, "confidence")]];
        case "probs": {
          const rows = Object.entries(answer.probabilities || {})
            .map(([k, p]) => [Number(k), Number(p)])
            .sort((a, b) => a[0] - b[0]);
          if (!rows.length) throw new JevError("Jev response is missing probabilities.", "na");
          return rows;
        }
      }
    }

    if (mode.type === "noul") {
      const p = requireNumber(answer.noul, "noul");
      switch (mode.output) {
        case "prob":
          return [[p]];
        case "bool":
          return [[p >= mode.threshold]];
        case "yesno":
          return [[p >= mode.threshold ? "Yes" : "No"]];
      }
    }

    throw new JevError(`Cannot format ${mode.type}:${mode.output}.`);
  }

  // ---------------------------------------------------------------------------
  // HTTP error mapping (mirrors the official SDK's message extraction)
  // ---------------------------------------------------------------------------

  function extractMessage(body) {
    if (typeof body === "string") return body.trim() || undefined;
    if (!body || typeof body !== "object") return undefined;
    const { error, message, detail } = body;
    if (typeof error === "string") return error;
    if (error && typeof error.message === "string") return error.message;
    if (typeof message === "string") return message;
    if (typeof detail === "string") return detail;
    if (detail && typeof detail.message === "string") return detail.message;
    if (Array.isArray(detail)) {
      return detail
        .map((d) => {
          const loc = Array.isArray(d.loc) ? d.loc.join(".") : "";
          return loc ? `${loc}: ${d.msg || d.message || ""}` : d.msg || d.message || JSON.stringify(d);
        })
        .join("; ");
    }
    return undefined;
  }

  function httpError(status, text) {
    let body = text;
    try {
      body = JSON.parse(text);
    } catch (_) {
      /* plain text */
    }
    const detail = extractMessage(body);
    const suffix = detail ? `: ${String(detail).slice(0, 200)}` : "";
    const prefixes = {
      400: "Bad request",
      401: "Invalid or missing API key — set it in the Jev pane",
      403: "Access denied for this API key",
      404: "Not found — check the base URL and model name",
      422: "Jev rejected the question",
      429: "Rate limited by Jev — try again shortly",
      529: "Jev is overloaded — try again shortly",
    };
    const prefix = prefixes[status] || (status >= 500 ? "Jev server error" : "Jev request failed");
    return new JevError(`${prefix} (HTTP ${status})${suffix}`, status >= 500 || status === 429 ? "na" : "value", status);
  }

  function parseRetryAfter(headers) {
    if (!headers || typeof headers.get !== "function") return undefined;
    const rawMs = headers.get("retry-after-ms");
    const ms = Number(rawMs);
    if (rawMs !== null && Number.isFinite(ms) && ms >= 0) return ms;
    const raw = headers.get("retry-after");
    if (raw === null || raw === undefined) return undefined;
    const secs = Number(raw);
    if (Number.isFinite(secs) && secs >= 0) return secs * 1000;
    const date = Date.parse(raw);
    if (Number.isFinite(date)) return Math.max(0, date - Date.now());
    return undefined;
  }

  function backoffMs(attempt) {
    const base = Math.min(500 * 2 ** attempt, 5000);
    return base - Math.random() * base * 0.25;
  }

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  // ---------------------------------------------------------------------------
  // Client: caching, in-flight de-duplication, batching and concurrency control
  // ---------------------------------------------------------------------------

  const SEP = "␟";

  // Every request needs the user's own key: proxies pass it through, never add one.
  const missingKeyError = () =>
    new JevError("No Jev API key set. Open the Jev pane (Home tab ▸ Jev) and paste your TypeSafe API key.");

  class JevClient {
    constructor(settings = {}) {
      this.settings = { ...DEFAULTS };
      this.configure(settings);
      this.cache = new Map(); // cacheKey -> Promise<answer>
      this.pending = new Map(); // batchKey -> { model, state, items[] }
      this.queue = [];
      this.active = 0;
      this.flushTimer = null;
      this.lastModel = null;
      this.resetStats();
    }

    configure(settings = {}) {
      for (const [k, v] of Object.entries(settings)) {
        if (!(k in DEFAULTS) || v === undefined) continue;
        this.settings[k] = v;
      }
      const s = this.settings;
      s.apiKey = String(s.apiKey || "").trim();
      s.baseUrl = String(s.baseUrl || DEFAULTS.baseUrl).trim().replace(/\/+$/, "");
      s.model = String(s.model || DEFAULTS.model).trim();
      s.cache = s.cache !== false;
      return this;
    }

    resetStats() {
      this.stats = { calls: 0, cacheHits: 0, requests: 0, questions: 0, inputTokens: 0, errors: 0 };
    }

    clearCache() {
      this.cache.clear();
    }

    usesOfficialApi() {
      return /^https:\/\/api\.typesafe\.ai$/i.test(this.settings.baseUrl);
    }

    /** Ask one question about one state. Resolves to the raw Jev answer object. */
    ask(state, question) {
      this.stats.calls++;
      if (!this.settings.apiKey) return Promise.reject(missingKeyError());
      const model = this.settings.model;
      const stateKey = JSON.stringify(state);
      const cacheKey = model + SEP + stateKey + SEP + JSON.stringify(question);

      if (this.settings.cache) {
        const hit = this.cache.get(cacheKey);
        if (hit) {
          this.stats.cacheHits++;
          // refresh LRU position
          this.cache.delete(cacheKey);
          this.cache.set(cacheKey, hit);
          return hit;
        }
      }

      this.stats.questions++;
      const promise = new Promise((resolve, reject) => {
        const batchKey = model + SEP + stateKey;
        let batch = this.pending.get(batchKey);
        if (!batch) {
          batch = { model, state, items: [] };
          this.pending.set(batchKey, batch);
        }
        batch.items.push({ question, resolve, reject });
        if (batch.items.length >= this.settings.maxBatchSize) {
          this.pending.delete(batchKey);
          this._enqueue(batch);
        } else {
          this._scheduleFlush();
        }
      });

      if (this.settings.cache) {
        this.cache.set(cacheKey, promise);
        this._trimCache();
        promise.catch(() => {
          if (this.cache.get(cacheKey) === promise) this.cache.delete(cacheKey);
        });
      }
      return promise;
    }

    /** GET /v1/models -> [{ name, description, release_date }] */
    async listModels() {
      if (!this.settings.apiKey) throw missingKeyError();
      const data = await this._request("GET", "/v1/models");
      const models = Array.isArray(data) ? data : data && data.models;
      if (!Array.isArray(models)) throw new JevError("Unexpected response from /v1/models.", "na");
      return models;
    }

    _trimCache() {
      const max = this.settings.maxCacheEntries;
      while (this.cache.size > max) {
        this.cache.delete(this.cache.keys().next().value);
      }
    }

    _scheduleFlush() {
      if (this.flushTimer) return;
      this.flushTimer = setTimeout(() => {
        this.flushTimer = null;
        const batches = [...this.pending.values()];
        this.pending.clear();
        batches.forEach((b) => this._enqueue(b));
      }, this.settings.batchWindowMs);
    }

    _enqueue(batch) {
      this.queue.push(batch);
      this._pump();
    }

    _pump() {
      while (this.active < this.settings.maxConcurrency && this.queue.length) {
        const batch = this.queue.shift();
        this.active++;
        this._send(batch).finally(() => {
          this.active--;
          this._pump();
        });
      }
    }

    async _send(batch) {
      const questions = {};
      batch.items.forEach((item, i) => {
        questions["q" + i] = item.question;
      });
      try {
        const data = await this._request("POST", "/v1/systemone", {
          model: batch.model,
          state: batch.state,
          questions,
        });
        if (data && data.usage) this.stats.inputTokens += Number(data.usage.input_tokens) || 0;
        if (data && data.model) this.lastModel = data.model;
        const answers = (data && data.answers) || {};
        batch.items.forEach((item, i) => {
          const a = answers["q" + i];
          if (a) item.resolve(a);
          else item.reject(new JevError("Jev returned no answer for this question.", "na"));
        });
      } catch (err) {
        // One bad question in a batch fails the whole request; retry individually so the
        // error lands only on the cell that caused it.
        if (err && err.status === 422 && batch.items.length > 1) {
          batch.items.forEach((item) => this._enqueue({ model: batch.model, state: batch.state, items: [item] }));
          return;
        }
        this.stats.errors++;
        batch.items.forEach((item) => item.reject(err));
      }
    }

    async _request(method, path, body) {
      const s = this.settings;
      const doFetch = s.fetch || (typeof fetch !== "undefined" ? fetch.bind(globalThis) : null);
      if (!doFetch) throw new JevError("No fetch implementation available.");
      const url = s.baseUrl + path;
      const headers = { Accept: "application/json" };
      if (body !== undefined) headers["Content-Type"] = "application/json";
      headers.Authorization = "Bearer " + s.apiKey;
      const payload = body !== undefined ? JSON.stringify(body) : undefined;

      for (let attempt = 0; ; attempt++) {
        this.stats.requests++;
        const ctrl = typeof AbortController !== "undefined" ? new AbortController() : null;
        const timer = ctrl ? setTimeout(() => ctrl.abort(), s.timeoutMs) : null;
        let res;
        let text;
        try {
          res = await doFetch(url, { method, headers, body: payload, signal: ctrl ? ctrl.signal : undefined });
          text = await res.text();
        } catch (err) {
          if (timer) clearTimeout(timer);
          if (attempt < s.maxRetries) {
            await sleep(backoffMs(attempt));
            continue;
          }
          const timedOut = err && err.name === "AbortError";
          let message;
          if (timedOut) message = `Jev request timed out after ${s.timeoutMs} ms.`;
          else if (this.usesOfficialApi())
            message =
              "Could not reach the Jev API (network or CORS error). api.typesafe.ai blocks calls from browser-based add-ins: clear Base URL in the Jev pane to use the add-in's own server.";
          else message = `Could not reach ${s.baseUrl} (network or CORS error). Check your connection and the Base URL.`;
          throw new JevError(message, "na");
        }
        if (timer) clearTimeout(timer);

        if (res.ok) {
          try {
            return text ? JSON.parse(text) : {};
          } catch (_) {
            throw new JevError("Jev returned a response that is not valid JSON.", "na");
          }
        }
        const retriable = res.status === 408 || res.status === 429 || res.status >= 500;
        if (retriable && attempt < s.maxRetries) {
          const hinted = parseRetryAfter(res.headers);
          await sleep(hinted !== undefined ? Math.min(hinted, 60000) : backoffMs(attempt));
          continue;
        }
        throw httpError(res.status, text);
      }
    }
  }

  // ---------------------------------------------------------------------------
  // One-shot helper used by the Excel functions
  // ---------------------------------------------------------------------------

  /** Evaluate JEV(mode, text, ...args) end-to-end. Returns a 2-D array for Excel. */
  async function evaluate(client, mode, text, args) {
    const m = parseMode(mode);
    const state = normalizeState(text);
    if (state === null) return [[""]]; // blank input -> blank output, no API call
    const { question, meta } = buildQuestion(m.type, args);
    const answer = await client.ask(state, question);
    return formatAnswer(answer, m, meta);
  }

  return {
    VERSION,
    DEFAULTS,
    JevError,
    JevClient,
    parseMode,
    normalizeState,
    buildStateJson,
    buildQuestion,
    formatAnswer,
    evaluate,
  };
});
