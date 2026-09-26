# Changelog

## 1.0.0 — 2026-09-24

First release.

- Office add-in (Windows, Mac, Excel on the web) with custom functions
  `JEV.ASK`, `JEV.CHOICE`, `JEV.SCORE`, `JEV.NOUL`, `JEV.STATE` and `JEV.MODELS`.
- All three Jev question types: `choice`, `score` and `noul`, with output selectors
  (`:confidence`, `:prob`, `:probs`, `:level`, `:label`, `:bool[:threshold]`, `:yesno`, `:json`).
- Batches questions about the same text into one request, caches answers, limits concurrency,
  and retries 408/429/5xx responses, honouring `Retry-After`.
- Settings task pane: API key, model, base URL, connection test, cache and usage stats.
- VBA edition (`vba/modJev.bas`) for Windows desktop Excel with the literal `=JEV("choice", …)` syntax.
- Hosting on a single Cloudflare Worker (`npm run deploy`). It serves the add-in, proxies Jev calls
  (`api.typesafe.ai` rejects browser requests) and serves a manifest pointing at its own URL.
  The local dev server proxies in the same way. Neither holds a key: each user's own
  TypeSafe key is passed through.
