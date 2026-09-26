/*
 * Jev for Excel — task pane (settings, connection test, cache/stats).
 * Shares the JavaScript runtime with functions.js via window.JevAddin.
 */
/* global Office, Excel, JevCore, window, document, performance */
(function () {
  "use strict";

  const $ = (id) => document.getElementById(id);

  function setStatus(text, kind) {
    const el = $("status");
    el.textContent = text;
    el.className = "status" + (kind ? " " + kind : "");
  }

  function fillForm() {
    const s = window.JevAddin.client.settings;
    $("apiKey").value = s.apiKey;
    $("model").value = s.model === JevCore.DEFAULTS.model ? "" : s.model;
    $("baseUrl").value = s.baseUrl === window.JevAddin.defaultBaseUrl ? "" : s.baseUrl;
    $("baseUrl").placeholder = window.JevAddin.defaultBaseUrl;
    $("cache").checked = s.cache;
  }

  function readForm() {
    return {
      apiKey: $("apiKey").value.trim(),
      model: $("model").value.trim() || JevCore.DEFAULTS.model,
      baseUrl: $("baseUrl").value.trim() || window.JevAddin.defaultBaseUrl,
      cache: $("cache").checked,
    };
  }

  function renderStats() {
    const { client } = window.JevAddin;
    const s = client.stats;
    const rows = [
      ["Model", client.lastModel || client.settings.model],
      ["Function calls", s.calls],
      ["Served from cache", s.cacheHits],
      ["Questions sent", s.questions],
      ["HTTP requests", s.requests],
      ["Input tokens", s.inputTokens.toLocaleString()],
      ["Errors", s.errors],
      ["Cached answers", client.cache.size],
    ];
    const items = rows.flatMap(([label, value]) => {
      const dt = document.createElement("dt");
      const dd = document.createElement("dd");
      dt.textContent = label;
      dd.textContent = String(value);
      return [dt, dd];
    });
    $("stats").replaceChildren(...items);
  }

  async function recalcAll() {
    await Excel.run(async (ctx) => {
      ctx.workbook.application.calculate(Excel.CalculationType.full);
      await ctx.sync();
    });
  }

  function wire() {
    $("version").textContent = "v" + JevCore.VERSION;
    fillForm();
    renderStats();
    setInterval(renderStats, 1500);

    $("toggleKey").addEventListener("click", () => {
      const input = $("apiKey");
      const show = input.type === "password";
      input.type = show ? "text" : "password";
      $("toggleKey").textContent = show ? "Hide" : "Show";
      $("toggleKey").setAttribute("aria-label", show ? "Hide key" : "Show key");
    });

    $("save").addEventListener("click", async () => {
      window.JevAddin.saveSettings(readForm());
      renderStats();
      setStatus("Saved. Recalculating…", "ok");
      try {
        await recalcAll();
        setStatus("Saved.", "ok");
      } catch (e) {
        setStatus("Saved. (Recalculate manually: " + e.message + ")", "ok");
      }
    });

    $("test").addEventListener("click", async () => {
      window.JevAddin.saveSettings(readForm());
      const btn = $("test");
      btn.disabled = true;
      setStatus("Testing…");
      try {
        const probe = new JevCore.JevClient({ ...window.JevAddin.client.settings, cache: false, maxRetries: 0 });
        const t0 = performance.now();
        const answer = await probe.ask("I was charged twice for my subscription this month.", {
          type: "noul",
          instructions: "The message is about billing or payments",
        });
        const ms = Math.round(performance.now() - t0);
        setStatus(`Connected ✓  ${probe.lastModel || ""} answered in ${ms} ms (p(yes) = ${Number(answer.noul).toFixed(3)})`, "ok");
      } catch (e) {
        setStatus(e.message || String(e), "err");
      } finally {
        btn.disabled = false;
        renderStats();
      }
    });

    $("recalc").addEventListener("click", async () => {
      try {
        await recalcAll();
        setStatus("Recalculated.", "ok");
      } catch (e) {
        setStatus(e.message, "err");
      }
    });

    $("clearCache").addEventListener("click", () => {
      window.JevAddin.client.clearCache();
      window.JevAddin.client.resetStats();
      renderStats();
      setStatus("Cache cleared. The next recalculation will call Jev again.", "ok");
    });
  }

  Office.onReady(() => wire());
})();
