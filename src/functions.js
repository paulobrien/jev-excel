/*
 * Jev for Excel — custom function implementations.
 *
 * Runs in the add-in's shared runtime (loaded by taskpane.html), so the
 * JevClient instance, its cache and its settings are shared with the task pane.
 */
/* global CustomFunctions, JevCore, window, localStorage */
(function () {
  "use strict";

  const STORAGE_KEY = "jev-excel.settings.v1";

  function loadSettings() {
    try {
      return JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}") || {};
    } catch (_) {
      return {};
    }
  }

  // api.typesafe.ai doesn't accept browser (CORS) requests, so the add-in calls the
  // server it was loaded from (the dev server or the Cloudflare Worker), which proxies
  // to TypeSafe. A stored api.typesafe.ai URL could never work here, so it's ignored.
  const defaultBaseUrl = window.location.origin;

  const initial = loadSettings();
  if (initial.baseUrl === JevCore.DEFAULTS.baseUrl) delete initial.baseUrl;
  const client = new JevCore.JevClient({ baseUrl: defaultBaseUrl, ...initial });

  function saveSettings(settings) {
    const merged = { ...loadSettings(), ...settings };
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(merged));
    } catch (_) {
      /* storage unavailable — settings still apply for this session */
    }
    client.configure(merged);
    return merged;
  }

  // Exposed to the task pane (same JavaScript runtime).
  window.JevAddin = { client, loadSettings, saveSettings, STORAGE_KEY, defaultBaseUrl };

  // ---------------------------------------------------------------------------

  function toExcelError(err) {
    const message = String((err && err.message) || err || "Unknown error").slice(0, 250);
    const code =
      err && err.kind === "na" ? CustomFunctions.ErrorCode.notAvailable : CustomFunctions.ErrorCode.invalidValue;
    return new CustomFunctions.Error(code, message);
  }

  function asArgs(repeating) {
    if (repeating === null || repeating === undefined) return [];
    return Array.isArray(repeating) ? repeating : [repeating];
  }

  function isBlank(arg) {
    return JevCore.normalizeState(arg) === null;
  }

  async function run(mode, text, args) {
    try {
      return await JevCore.evaluate(client, mode, text, args);
    } catch (err) {
      throw toExcelError(err);
    }
  }

  /** =JEV.ASK(mode, text, [arg1], [arg2], ...) */
  function ask(mode, text, options) {
    return run(mode, text, asArgs(options));
  }

  /** =JEV.CHOICE(text, option1, option2, ...) */
  function choice(text, options) {
    return run("choice", text, asArgs(options));
  }

  /** =JEV.SCORE(text, level0, level1, ...) */
  function score(text, levels) {
    return run("score", text, asArgs(levels));
  }

  /** =JEV.NOUL(text, question, [whenTrue], [whenFalse]) */
  function noul(text, question, whenTrue, whenFalse) {
    // The core reads the descriptions by position, so a FALSE description on its
    // own would be taken as the TRUE one.
    if (isBlank(whenTrue) && !isBlank(whenFalse)) {
      return Promise.reject(toExcelError(new JevCore.JevError("Give a when_true description before when_false.")));
    }
    return run("noul", text, [question, whenTrue, whenFalse]);
  }

  /** =JEV.STATE(headers, values) -> JSON object text */
  function state(headers, values) {
    try {
      return JevCore.buildStateJson(headers, values);
    } catch (err) {
      throw toExcelError(err);
    }
  }

  /** =JEV.MODELS() -> spill of name | description | release date */
  async function models() {
    try {
      const list = await client.listModels();
      if (!list.length) return [["(no models)", "", ""]];
      return list.map((m) => [
        String(m.name ?? m.id ?? ""),
        String(m.description ?? ""),
        String(m.release_date ?? ""),
      ]);
    } catch (err) {
      throw toExcelError(err);
    }
  }

  if (typeof CustomFunctions !== "undefined") {
    CustomFunctions.associate("ASK", ask);
    CustomFunctions.associate("CHOICE", choice);
    CustomFunctions.associate("SCORE", score);
    CustomFunctions.associate("NOUL", noul);
    CustomFunctions.associate("STATE", state);
    CustomFunctions.associate("MODELS", models);
  }
})();
