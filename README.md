# Jev for Excel

Call [TypeSafe AI's Jev](https://docs.typesafe.ai/) from a worksheet formula.

```excel
=JEV.ASK("choice", A2, "billing", "technical", "sales", "other")      → technical
=JEV.ASK("score",  A2, "Calm", "Frustrated but civil", "Very angry")   → 1.43
=JEV.ASK("noul",   A2, "The customer is asking for a refund")          → 0.93
```

Jev is a "System One" model. It doesn't write text. It gives you a typed, calibrated decision: which option fits, where something sits on a scale, or how likely a yes/no statement is to be true. That's the kind of answer you want in a spreadsheet cell, where you can filter, sort, pivot and chart it. Fill a formula down a column of support tickets, product reviews, survey responses or emails, and each row gets classified in milliseconds.

> **Unofficial.** This project isn't affiliated with or endorsed by TypeSafe AI. You need your own TypeSafe API key.

---

## Contents

- [Two editions](#two-editions)
- [Quick start](#quick-start)
- [Function reference](#function-reference)
  - [JEV.ASK: the main function](#jevask-the-main-function)
  - [Modes and outputs](#modes-and-outputs)
  - [Options, rubric levels and questions](#options-rubric-levels-and-questions)
  - [Shortcut functions](#shortcut-functions)
  - [JEV.STATE: evaluating several fields](#jevstate-evaluating-several-fields)
  - [JEV.MODELS](#jevmodels)
  - [Errors](#errors)
- [Worked examples](#worked-examples)
- [How calls are made (batching, caching, cost)](#how-calls-are-made)
- [Installation](#installation)
  - [1. Run it locally](#1-run-it-locally)
  - [2. Deploy to Cloudflare](#2-deploy-to-cloudflare)
  - [3. Add it to Excel](#3-add-it-to-excel)
- [Settings and security](#settings-and-security)
- [Proxy](#proxy)
- [VBA edition](#vba-edition)
- [Limits and tips](#limits-and-tips)
- [Troubleshooting](#troubleshooting)
- [Development](#development)
- [Licence](#licence)

---

## Two editions

| | **Office add-in** (recommended) | **VBA edition** |
|---|---|---|
| Formula | `=JEV.ASK("choice", A2, "opt1", "opt2")` | `=JEV("choice", A2, "opt1", "opt2")` |
| Runs on | Excel for Windows, Mac and the web (Microsoft 365) | Windows desktop Excel only |
| Calls | Asynchronous, batched and parallel. Excel stays responsive. | One at a time. Excel waits for each cell. |
| Install | Sideload a manifest or deploy centrally from the M365 admin centre | Import one `.bas` file and save it as an `.xlam` |
| Code | [`src/`](src) | [`vba/modJev.bas`](vba/modJev.bas) |

Office add-ins always put custom functions under a namespace, so the add-in's functions are `JEV.ASK`, `JEV.CHOICE` and so on. If you want the bare `=JEV(...)` spelling, use the VBA edition. Both take the same arguments and give the same outputs.

---

## Quick start

1. **Get an API key** from TypeSafe AI (see their [API docs](https://docs.typesafe.ai/api)).
2. **Install the add-in.**
   - **To try it on your own PC,** run `npm install`, `npm run certs` and `npm start`. Excel opens with the add-in loaded ([details](#1-run-it-locally)).
   - **For day-to-day use or sharing,** [deploy it to Cloudflare](#2-deploy-to-cloudflare) and [add it to Excel](#3-add-it-to-excel).
   - **On Excel 2016/2019,** use the [VBA edition](#vba-edition).
3. In Excel, open **Home ▸ Jev**, paste your key, and click **Test connection**.
4. Put some text in `A2` and try:
   ```excel
   =JEV.ASK("choice", A2, "complaint", "question", "praise")
   ```
5. Fill it down the column.

---

## Function reference

For a one-page summary, see the **cheat sheet**. It's [`src/cheatsheet.html`](src/cheatsheet.html) in the repo, and the add-in's server hosts it at `/cheatsheet.html`, for example `https://jev-excel.<your-subdomain>.workers.dev/cheatsheet.html`. The task pane links to it too. It's a single self-contained file, so you can also download it, open it in any browser or print it.

### JEV.ASK: the main function

```
=JEV.ASK(mode, text, [arg1], [arg2], ...)
```

| Argument | What it is |
|---|---|
| `mode` | `"choice"`, `"score"` or `"noul"`, optionally followed by an output selector such as `"choice:probs"` or `"noul:bool:0.7"`. See [Modes and outputs](#modes-and-outputs). |
| `text` | The text Jev should evaluate: a string, a cell or a range. A blank cell returns a blank result and makes no API call. A multi-cell range is sent as a list of its non-empty values. A cell that holds valid JSON (an object or array) is sent as structured data. |
| `arg1…` | Choice options, score rubric levels, or the yes/no question. Each one can be text or a range, and you can pass up to 253. Any argument starting with `?` is used as the question. See [below](#options-rubric-levels-and-questions). |

### Modes and outputs

The mode string is `type[:output[:threshold]]`. It isn't case-sensitive.

#### `choice`: pick one of N options

| Mode | Returns |
|---|---|
| `choice` (or `choice:label`) | The chosen label, e.g. `billing` |
| `choice:confidence` | Jev's confidence in that choice (0–1) |
| `choice:prob` | The probability of the chosen label (0–1) |
| `choice:probs` | A **spilled 2-column table** of every label and its probability, highest first |
| `choice:json` | The raw answer as JSON text |

Aliases for `choice`: `classify`, `category`, `choose`, `pick`, `select`.

#### `score`: place the text on an ordered rubric

| Mode | Returns |
|---|---|
| `score` | The expected level, a probability-weighted value that can fall between levels (e.g. `1.43`) |
| `score:level` | The nearest whole level (`0`, `1`, `2`, …) |
| `score:label` | The rubric text for the nearest level |
| `score:confidence` | Jev's confidence (0–1) |
| `score:probs` | A **spilled 2-column table** of level and probability |
| `score:json` | The raw answer as JSON text |

Aliases for `score`: `rate`, `rating`, `scale`, `rubric`.

#### `noul`: yes/no probability

| Mode | Returns |
|---|---|
| `noul` | The probability that the statement is true (0–1) |
| `noul:bool` | `TRUE` if the probability is at least 0.5, otherwise `FALSE` |
| `noul:bool:0.8` | `TRUE`/`FALSE` using your own threshold |
| `noul:yesno` / `noul:yesno:0.8` | `"Yes"` / `"No"` |
| `noul:json` | The raw answer as JSON text |

Aliases for `noul`: `yesno`, `yn`, `bool`, `boolean`, `binary`.

### Options, rubric levels and questions

**Choice options.** Each argument after `text` is a label. You can pass them inline or as a range.

```excel
=JEV.ASK("choice", A2, "billing", "technical", "other")
=JEV.ASK("choice", A2, $F$2:$F$9)                        ← labels listed in F2:F9
=JEV.ASK("choice", A2, $B$1:$E$1)                        ← labels across a header row
```

Labels become the values Jev returns, so keep them short and stable. Descriptions help Jev tell similar options apart. You can add them in two ways:

```excel
=JEV.ASK("choice", A2, "billing | Payments, invoices, refunds", "technical | Bugs or outages", "other")
=JEV.ASK("choice", A2, $F$2:$G$9)     ← 2 columns and 2+ rows: label in F, description in G
```

Jev supports 2–255 options per question. TypeSafe recommends including an explicit `other` option.

**Score levels.** List the levels from lowest (level 0) to highest. You need 2–10.

```excel
=JEV.ASK("score", A2, "Cosmetic, no impact", "Degraded but has a workaround", "Blocking, no workaround")
=JEV.ASK("score", A2, $H$2:$H$5)
```

**The question (`?` prefix).** `choice` and `score` work without a question. Jev infers the task from the options. When the task isn't obvious from the options, add a question by starting any argument with `?`:

```excel
=JEV.ASK("choice", A2, "?Which team should handle this ticket?", "billing", "technical", "sales")
=JEV.ASK("score",  A2, "?How urgent is this request?", "Routine", "Soon", "Today", "Immediately")
```

**Noul (yes/no).** The first argument is the question or statement. You can add descriptions of what counts as TRUE and FALSE after it.

```excel
=JEV.ASK("noul", A2, "The customer is asking for a refund")
=JEV.ASK("noul", A2, "Is this urgent?", "A deadline or outage is mentioned", "No time pressure")
```

Statements often work as well as questions. Instructions can refer to fields in structured state using `backticks` (see [JEV.STATE](#jevstate-evaluating-several-fields)).

### Shortcut functions

These do the same thing with less typing:

| Shortcut | Same as |
|---|---|
| `=JEV.CHOICE(text, option1, option2, …)` | `=JEV.ASK("choice", text, option1, option2, …)` |
| `=JEV.SCORE(text, level0, level1, …)` | `=JEV.ASK("score", text, level0, level1, …)` |
| `=JEV.NOUL(text, question, [when_true], [when_false])` | `=JEV.ASK("noul", text, question, when_true, when_false)` |

The shortcuts return the default output for their type. For any other output, use `JEV.ASK` with a mode string. Identical questions share one cached answer, so `=JEV.CHOICE(A2, …)` and `=JEV.ASK("choice:probs", A2, …)` with the same options make only one API call between them.

### JEV.STATE: evaluating several fields

Jev can evaluate a JSON object of named fields instead of plain text. `JEV.STATE` builds that object from a header row and a data row:

```excel
=JEV.STATE($A$1:$D$1, A2:D2)
→ {"Order":"A-1042","Product":"Headphones","Review":"Left earcup stopped working after a week…","Stars":2}
```

Pass it as the `text` argument, and refer to fields in your question with backticks:

```excel
=JEV.ASK("noul", JEV.STATE($A$1:$D$1, A2:D2), "`Review` describes a product defect")
```

The VBA edition calls this `JEV_STATE`.

### JEV.MODELS

`=JEV.MODELS()` spills a list of the models your key can use, with columns for name, description and release date. Set the model in the task pane. You can pin a version such as `jev-1.13.0` so that probabilities and thresholds stay stable when `jev-latest` moves on.

### Errors

| You see | Meaning |
|---|---|
| `#VALUE!` | Bad arguments, e.g. an unknown mode, fewer than 2 options, a missing API key or a 401/422 from the API. **Hover over the cell** (or select it and click the error icon) to read the message. |
| `#N/A` | Jev couldn't be reached, timed out, was rate-limited or overloaded after retries, or returned an unexpected response. Recalculate to try again. |
| `#BUSY!` | The request is still in flight. |

In the VBA edition, errors appear as text like `#JEV! No Jev API key…`. Run the `JevToggleErrorStyle` macro to switch to `#VALUE!`.

### Unquoted modes: `=JEV.ASK(choice, …)`

If you'd rather not type quotes, define three workbook names under **Formulas ▸ Name Manager ▸ New**: `choice` = `="choice"`, `score` = `="score"` and `noul` = `="noul"`. Then `=JEV.ASK(choice, A2, "a", "b")` works. In the VBA edition, the `JevAddModeNames` macro adds these names for you. The VBA edition also tries to read an unquoted mode straight from the formula text, as a fallback.

---

## Worked examples

**Support ticket triage** (ticket text in column A):

| | Formula | Result |
|---|---|---|
| Team | `=JEV.CHOICE(A2, "?Which team should handle this?", "billing", "technical", "account", "other")` | `billing` |
| How sure | `=JEV.ASK("choice:confidence", A2, "?Which team should handle this?", "billing", "technical", "account", "other")` | `0.62` |
| Anger (0–2) | `=JEV.SCORE(A2, "Calm, just stating facts", "Frustrated but civil", "Very angry, strong language")` | `1.12` |
| Escalate? | `=JEV.ASK("noul:bool:0.7", A2, "The customer threatens to cancel or complain publicly")` | `FALSE` |
| All probabilities | `=JEV.ASK("choice:probs", A2, $F$2:$F$5)` | spills a 4×2 table |

The first two formulas ask the same question about the same text, so they share one cached answer. The anger score and the escalation check are about the same text as well, so they go out batched in the same HTTP request.

**Product reviews** (review text in column C):

```excel
=JEV.CHOICE(C2, "?What is the review mainly about?", $K$2:$L$8)     ← topics + descriptions
=JEV.ASK("score:label", C2, "?Overall sentiment", "Very negative", "Negative", "Mixed", "Positive", "Very positive")
=JEV.NOUL(C2, "The reviewer reports a safety problem")
=JEV.NOUL(C2, "The review looks fake or incentivised", "Generic praise, no product detail, mentions a free sample", "Specific first-hand experience")
```

**Rules on top of probabilities.** The outputs are plain numbers, so ordinary Excel logic works on them:

```excel
=IF(JEV.NOUL(A2, "Mentions a legal deadline") > 0.8, "Legal", JEV.CHOICE(A2, $F$2:$F$6))
=COUNTIFS(Results!D:D, ">=0.7")
```

---

## How calls are made

Every call becomes a `POST /v1/systemone` request. The request goes to the server the add-in was loaded from, which forwards it to `api.typesafe.ai` (see [Proxy](#proxy)):

```json
{
  "model": "jev-latest",
  "state": "I was charged twice for my subscription.",
  "questions": {
    "q0": { "type": "choice", "instructions": "Which team?", "criteria": { "billing": null, "technical": null, "other": null } }
  }
}
```

Several things keep this fast and cheap:

- **Batching.** Calls are gathered for about 40 ms. Questions about the same text (and model) go out together in one request, up to 16 at a time. Five formulas across a row of ticket text cost one round trip.
- **Caching.** Answers are cached in memory for the Excel session, keyed by model, text and question. Recalculating with F9, reopening the task pane, or asking for a different output from the same question doesn't call the API again. Errors are never cached. You can clear the cache from the task pane.
- **Concurrency.** At most 6 requests are in flight at once, which keeps large fill-downs within Jev's documented limit of 1,200 requests a minute.
- **Retries.** Responses with status 408, 429 or 5xx (including 529 "overloaded") and network failures are retried up to 3 times with exponential backoff. The add-in honours `Retry-After` and `retry-after-ms`.
- **Blank in, blank out.** An empty `text` cell returns `""` without making a call.
- **Isolated validation errors.** If one question in a batch is invalid (HTTP 422), the batch is retried one question at a time, so only the offending cell shows an error.

TypeSafe bills input tokens only; output is free at the time of writing. Check their pricing page for current rates. The task pane shows how many questions you've sent, how many input tokens they used, and your cache hits.

> **Freezing results.** Custom functions recalculate when their inputs change, and Excel may recalculate them when a workbook opens. To lock in a set of answers, copy the column and use **Paste ▸ Values**.

---

## Installation

An Office add-in has two parts:

- **The add-in's files** (`src/`: HTML, JS, JSON and icons). They have to be served over HTTPS. While you develop, a local server on your PC serves them. For real use, a Cloudflare Worker does.
- **A manifest** (`manifest.xml`). This small XML file tells Excel the add-in's name and where its files are. You add the add-in to Excel by giving Excel a manifest.

Either way, the server also forwards Jev calls to TypeSafe, because TypeSafe doesn't accept calls straight from an add-in (see [Proxy](#proxy)). Every user pastes their own TypeSafe key into the add-in.

**Pick a route:**

| You want to… | Do this |
|---|---|
| Try it or work on it on your own PC | [1. Run it locally](#1-run-it-locally). On desktop Excel, `npm start` also adds it to Excel for you. |
| Use it day to day, or share it | [2. Deploy to Cloudflare](#2-deploy-to-cloudflare), then [3. Add it to Excel](#3-add-it-to-excel) |
| Use Excel 2016/2019, or the bare `=JEV(...)` syntax | The [VBA edition](#vba-edition). No Node, server or manifest needed. |

**Requirements:** Excel with the SharedRuntime 1.1 requirement set. That means Microsoft 365 on Windows or Mac, or Excel on the web. Perpetual Excel 2016/2019 isn't supported, so use the VBA edition there. To run or deploy the add-in you also need [Node.js](https://nodejs.org/) 20 or later.

### 1. Run it locally

The local version runs from a dev server on your PC at `https://localhost:3000`. That server serves the add-in's files and forwards Jev calls to TypeSafe. Excel loads the add-in from it using the repo's own [`manifest.xml`](manifest.xml), which points at `localhost`. Only the PC running the server can use this version.

**One-off setup,** in the repo folder:

```bash
npm install
```
```bash
npm run certs
```

`npm run certs` creates an HTTPS certificate for `localhost` and asks Windows or macOS to trust it. Click **Yes** when the security prompt asks whether to install the certificate. Office only loads add-ins over HTTPS.

#### Desktop Excel (Windows or Mac): automatic

```bash
npm start
```

This does three things:

- It opens a second terminal window running the dev server. **Leave that window open.** If it closes, the add-in stops working and cells show errors.
- It registers `manifest.xml` with desktop Excel as a developer add-in. On a Mac, it copies the manifest into Excel's `wef` folder.
- It opens Excel with a new workbook that has the add-in loaded.

Then:

1. In Excel, look at the right-hand end of the **Home** tab for a **Jev** group with a **Jev** button. Click it to open the settings pane.
2. Paste your TypeSafe API key, click **Save**, then click **Test connection**.
3. In a cell, type `=JEV.`. Excel's autocomplete should list `JEV.ASK`, `JEV.CHOICE` and the rest.

**Next time:** the add-in stays registered and loads in every new Excel window. Only the server needs restarting. Run this and leave it running before you open Excel:

```bash
npm run serve
```

If you opened Excel without the server running, close Excel, start the server, then reopen Excel.

**When you're finished:** this unregisters the add-in and stops the server:

```bash
npm run stop
```

#### Excel on the web

`npm run start:web` starts the server and opens a workbook with the add-in loaded. It needs the address of a workbook in OneDrive or SharePoint. Open one in your browser, copy the address from the address bar, and pass it in:

```bash
npm run start:web -- --document "https://onedrive.live.com/edit.aspx?resid=..."
```

Alternatively, start the server yourself and add the repo's `manifest.xml` by hand, as described in [3. Add it to Excel](#3-add-it-to-excel):

```bash
npm run serve
```

Excel on the web loads the files from `localhost` in your browser, so this only works on the PC running the server.

#### Adding it by hand

`npm start` covers most cases. If you'd rather add the add-in yourself, for example on a PC where `npm start` can't open Excel:

1. Run `npm run serve` and leave it running.
2. Open `https://localhost:3000/taskpane.html` in a browser. You should see the settings page with no certificate warning. The page does nothing outside Excel, but if it loads, the server and certificate are working.
3. Follow [3. Add it to Excel](#3-add-it-to-excel) using the repo's `manifest.xml`.

### 2. Deploy to Cloudflare

For real use, deploy the add-in to [Cloudflare Workers](https://workers.cloudflare.com/). The free plan is enough. One Worker, named `jev-excel`, does everything from one URL:

- It serves the add-in's files.
- It forwards Jev calls to TypeSafe.
- It serves a `manifest.xml` that points at its own URL, and a `functions.json` whose **Help on this function** links open the right section of its cheat sheet.

It holds no keys and needs no secrets or settings. The code is [`worker/index.mjs`](worker/index.mjs), and it's configured in [`wrangler.toml`](wrangler.toml).

1. **Create a Cloudflare account** at [dash.cloudflare.com](https://dash.cloudflare.com/sign-up) if you don't have one.
2. **Log in from the repo folder.** This opens a browser so you can authorise Wrangler, Cloudflare's command-line tool. You only need to do it once per PC.
   ```bash
   npx wrangler@4 login
   ```
3. **Deploy.** This builds `dist/` from `src/` and uploads it together with the Worker:
   ```bash
   npm run deploy
   ```
   The first time, Wrangler may ask you to choose a `workers.dev` subdomain for your account.
4. **Note the URL** Wrangler prints, for example `https://jev-excel.<your-subdomain>.workers.dev`.
5. **Check it.** Open `https://jev-excel.<your-subdomain>.workers.dev/taskpane.html` in a browser. You should see the settings page. The manifest you'll give Excel is at `https://jev-excel.<your-subdomain>.workers.dev/manifest.xml`.

The manifest is generated for whatever address it's requested from, so it keeps working if you later [add a custom domain](https://developers.cloudflare.com/workers/configuration/routing/custom-domains/) to the Worker.

**Try the Worker locally first (optional).** This runs the same Worker code on your PC, at `http://localhost:8787`:

```bash
npm run build
```
```bash
npx wrangler@4 dev
```

It's plain HTTP, so Excel won't load the add-in from it, but you can check the routes in a browser or with `curl`. To work on the add-in itself, use [1. Run it locally](#1-run-it-locally).

**To update** the deployed add-in, run `npm run deploy` again.

### 3. Add it to Excel

This works the same for both versions. Only the manifest differs:

| Version | Manifest | Must be running |
|---|---|---|
| Local | [`manifest.xml`](manifest.xml) in the repo, which points at `https://localhost:3000` | `npm run serve` on the same PC |
| Deployed | `https://jev-excel.<your-subdomain>.workers.dev/manifest.xml`. Open it in a browser, press **Ctrl+S** (**⌘S** on a Mac) and save it as `manifest.xml`. | Nothing. The Worker is always on. |

Don't add both versions at once. They share the same add-in ID, so Excel treats them as the same add-in. [Remove](#updating-and-removing) one before adding the other.

#### Excel on the web

1. Open any workbook in Excel on the web.
2. Go to **Home ▸ Add-ins ▸ More Add-ins**. In some versions the path is **Insert ▸ Add-ins ▸ More Add-ins**.
3. Open the **My Add-ins** tab and click **Upload My Add-in** at the top right.
4. Click **Browse…**, choose your `manifest.xml` and click **Upload**.

The **Jev** button appears on the **Home** tab. The upload is stored in your browser, so you may need to repeat it on another browser or after clearing browser data.

#### Windows desktop

Desktop Excel for Windows won't open a manifest file directly. It reads manifests from a shared folder that you mark as trusted. The folder can be on your own PC.

1. **Create the folder.** Make a folder, for example `C:\JevManifest`, and put your `manifest.xml` in it.
2. **Share it.** Right-click the folder and choose **Properties ▸ Sharing ▸ Share…**. Click **Share**, then **Done**. Note the network path it shows, for example `\\YOUR-PC\JevManifest`. Excel needs this `\\…` path. A `C:\…` path won't work.
3. **Trust it in Excel.** Go to **File ▸ Options ▸ Trust Center ▸ Trust Center Settings… ▸ Trusted Add-in Catalogs**.
   1. Paste the `\\YOUR-PC\JevManifest` path into **Catalog Url** and click **Add catalog**.
   2. Tick **Show in Menu** for that row.
   3. Click **OK** twice.
4. **Restart Excel.** Close every Excel window and reopen it.
5. **Add it.** Go to **Home ▸ Add-ins ▸ More Add-ins**, or **Insert ▸ Get Add-ins ▸ My Add-ins** on older builds. Open the **SHARED FOLDER** tab, select **Jev for Excel** and click **Add**.

The **Jev** button appears on the **Home** tab, and Excel remembers the add-in from then on. To share the deployed version with colleagues, put its manifest on a network share everyone can read, and have each person do steps 3–5.

#### Mac

1. In Finder, choose **Go ▸ Go to Folder…** and enter `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`. If the folder doesn't exist, create the `wef` folder inside `Documents`.
2. Copy your `manifest.xml` into it.
3. Quit Excel completely (**⌘Q**) and reopen it.
4. Go to **Home ▸ Add-ins**, or **Insert ▸ My Add-ins**. Choose **Jev for Excel** from the list under **Developer Add-ins**.

#### Whole organisation (Microsoft 365 admin, deployed version only)

1. In the Microsoft 365 admin centre, go to **Settings ▸ Integrated apps ▸ Upload custom apps**.
2. Choose **Office Add-in**, then give the Worker's manifest URL or upload the saved `manifest.xml`.
3. Assign it to users or groups.

The add-in then appears on those users' **Home** tab without any further steps. It can take several hours to reach everyone.

#### Then, in Excel

1. Open **Home ▸ Jev**.
2. Paste your own TypeSafe API key and click **Save**. Leave **Base URL** blank, so the add-in uses the server it was loaded from.
3. Click **Test connection**.

### Check it's working

- The **Home** tab has a **Jev** button, and clicking it opens the settings pane.
- Typing `=JEV.` in a cell brings up the function list. If Excel shows `#NAME?` instead, the add-in isn't loaded in that window.
- After you save your key, **Test connection** reports success.

### Updating and removing

- **Updating the local version.** Changes to `src/` take effect the next time the task pane or Excel reloads. The dev server doesn't cache.
- **Updating the deployed version.** Run `npm run deploy`. People pick up the new version the next time they start Excel.
- **If Excel keeps showing an old version,** close Excel, clear the Office add-in cache, and reopen Excel. The cache is in `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` on Windows and `~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/` on a Mac.
- **Removing it.**
  - Added with `npm start`: run `npm run stop`.
  - Added by hand: go to **Home ▸ Add-ins ▸ More Add-ins ▸ My Add-ins**, click the **…** menu on **Jev for Excel** and choose **Remove**. On a Mac, also delete the manifest from the `wef` folder.
- **Taking the deployed version down.** Delete the `jev-excel` Worker in the Cloudflare dashboard, or run `npx wrangler@4 delete`.

---

## Settings and security

The **Jev** task pane (**Home ▸ Jev**) has these settings:

| Setting | Default | Notes |
|---|---|---|
| API key | none | Your own TypeSafe key. It's always needed. |
| Model | `jev-latest` | Pin a version, e.g. `jev-1.13.0`, if thresholds matter. |
| Base URL | The server the add-in was loaded from | Leave it blank. The dev server and the Cloudflare Worker both forward calls to TypeSafe (see [Proxy](#proxy)). Set it only to use a different proxy. |
| Cache answers | on | Turn off to force every recalculation to call Jev. |

The page also has **Test connection**, **Recalculate all**, **Clear cache** and live usage statistics.

**Where the key is stored:** in the add-in's own browser storage (`localStorage` for the add-in's origin) on that device, for that user. It never goes into the workbook, so you can share workbooks safely, but it isn't encrypted at rest. Anyone who uses your Windows or Mac profile can read it.

**What's sent:** only the `text` argument (or the `JEV.STATE` fields) and your question go to the API. Don't point Jev at data that your organisation's policies don't allow you to send to a third-party API.

---

## Proxy

**Why there's a proxy.** `api.typesafe.ai` rejects requests from browsers: its CORS check answers `Disallowed CORS origin` for any origin. Office add-ins run inside a browser engine, so they can't call it directly. Instead, the add-in calls `/v1/systemone` and `/v1/models` on the server it was loaded from, and that server forwards the calls to TypeSafe:

- **Locally**, that's the dev server behind `npm start` and `npm run serve`.
- **Deployed**, it's the Cloudflare Worker (see [Deploy to Cloudflare](#2-deploy-to-cloudflare)).

The add-in and its proxy share an origin, so there's no CORS to configure. The VBA edition isn't a browser, so it calls TypeSafe directly.

**The proxy never holds a key.** Each request carries the user's own TypeSafe key, which the proxy passes through unchanged. Usage is billed to that user's TypeSafe account, and the Worker's code doesn't log or store anything. That makes it safe to host one copy for anyone to use. Anyone who'd rather not send their key through your Worker can deploy their own copy from this repo.

To check a deployed Worker, replace the URL and key with your own:

```bash
curl -H "Authorization: Bearer YOUR_TYPESAFE_KEY" https://jev-excel.YOUR-SUBDOMAIN.workers.dev/v1/models
```

---

## VBA edition

For Windows desktop Excel, with the literal `=JEV(...)` syntax:

```excel
=JEV("choice", A2, "option1", "option2")
=JEV("score:label", A2, "Low", "Medium", "High")
=JEV("noul:bool", A2, "Mentions a deadline")
=JEV_STATE($A$1:$C$1, A2:C2)
=JEV_MODELS()
```

**Install:**

You need Windows desktop Excel 2010 or later (32- or 64-bit). You don't need Node, a web host or a manifest. You build a `Jev.xlam` add-in file once, then switch it on in Excel.

1. **Get the code.** You need [`vba/modJev.bas`](vba/modJev.bas), plus [`vba/ThisWorkbook.cls.txt`](vba/ThisWorkbook.cls.txt) if you want function help text (step 4). If you don't have a clone, open each file on GitHub and click **Download raw file**.
2. **Open the VBA editor.** In Excel, create a new blank workbook and press **Alt+F11**.
3. **Import the module.** In the VBA editor, choose **File ▸ Import File…** and pick `modJev.bas`. In the **Project** pane on the left, **Modules ▸ modJev** should now appear under **VBAProject (Book1)**. If the pane is hidden, press **Ctrl+R**.
4. *(Optional.)* **Add help text for the Insert Function dialog.** In the same project, open **Microsoft Excel Objects** and double-click **ThisWorkbook**. Paste the whole of `ThisWorkbook.cls.txt` into the code window that opens.
5. **Check it compiles.** Choose **Debug ▸ Compile VBAProject**. If nothing happens, it compiled without errors.
6. **Save it as an add-in.**
   1. Close the VBA editor and go back to Excel.
   2. Choose **File ▸ Save As ▸ Browse** and set **Save as type** to **Excel Add-in (\*.xlam)**. Excel switches to your add-ins folder (`%APPDATA%\Microsoft\AddIns`). Keep that location.
   3. Name the file `Jev` and click **Save**.
   4. Close the workbook. You don't need to save `Book1`.
7. **Switch it on.**
   1. Go to **File ▸ Options ▸ Add-ins**.
   2. At the bottom, set **Manage** to **Excel Add-ins** and click **Go…**.
   3. Tick **Jev** and click **OK**. If **Jev** isn't listed, click **Browse…** and pick `Jev.xlam`.

   Excel loads the add-in every time it starts from now on.
8. **Set your API key.**
   1. Press **Alt+F8**.
   2. Type `JevSetApiKey` in the **Macro name** box and click **Run**. The macro isn't in the list because add-in macros are hidden, but typing its name works.
   3. Paste your key when prompted.

   The key is stored in your Windows user registry. Instead of using the macro, you can set a `TYPESAFE_API_KEY` user environment variable and restart Excel.
9. **Try it.** In any workbook, enter:
   ```excel
   =JEV("noul", "I'd like my money back please", "The customer wants a refund")
   ```
   You should get a probability close to 1. `=JEV_MODELS()` is another quick connectivity check.

**Sharing it with colleagues.** Give them your `Jev.xlam`. Each person copies it into `%APPDATA%\Microsoft\AddIns` and follows steps 7–9. If the file came by email or download, Windows marks it as coming from the internet, and Excel blocks its macros. To prevent that, right-click the file, choose **Properties**, tick **Unblock** and click **OK**, all before loading it.

**Updating.**

1. Press **Alt+F11**.
2. In the Project pane, find **VBAProject (Jev.xlam)**, right-click **modJev** and choose **Remove modJev**. Answer **No** to exporting it.
3. Import the new `modJev.bas` into the same project.
4. With that project selected, press **Ctrl+S**.

**Removing.** Untick **Jev** in **File ▸ Options ▸ Add-ins ▸ Excel Add-ins ▸ Go…**, then delete `Jev.xlam`.

**Macros:** `JevSetApiKey`, `JevSetModel`, `JevSetBaseUrl`, `JevClearCache`, `JevShowStats`, `JevAddModeNames`, `JevToggleErrorStyle`, `JevRegisterFunctions`.

**Differences from the add-in:**

- Calls run one at a time and Excel waits for each. Jev usually answers in well under a second, but a fill-down over thousands of rows takes a while, so set calculation to **Manual** (Formulas ▸ Calculation Options) for large sheets.
- There's no batching. The session cache does apply, and it's cleared when Excel closes.
- Settings are stored in `HKCU\Software\VB and VBA Program Settings\JevForExcel`.
- HTTP goes through `MSXML2.XMLHTTP.6.0`, which respects the Windows system proxy. Change `HTTP_CLASS` at the top of the module if your network needs `WinHttp`.
- Errors appear as `#JEV! message` text by default, so you can see what went wrong.
- After an invalid key (401/403), or when Jev can't be reached after retries, the remaining cells fail immediately with the same message for 30 seconds instead of each one retrying. Changing the key or base URL, or running `JevClearCache`, clears this.
- The base URL must use `https://`. `http://` is accepted only for `localhost`, so the key never travels unencrypted.

Mac VBA can't use these Windows HTTP components. On a Mac, use the Office add-in.

---

## Limits and tips

These are Jev's documented limits (check TypeSafe's docs for the latest):

| Limit | Value |
|---|---|
| Choice options | 2–255 |
| Score rubric levels | 2–10 |
| State plus the longest question | about 32k tokens (roughly 150k characters) |
| Total per request | about 64k tokens |
| Rate limit | 1,200 requests a minute, 250k tokens a second |
| Input | text or JSON only |

**Tips:**

- **Word questions as statements.** "The customer is asking for a refund" often works better than "Refund?".
- **Keep each question to one thing.** Jev is weak at multi-part questions, arithmetic and indirect reasoning. Split a compound question into several columns and combine them with `AND`/`IF`.
- **Give options descriptions** when they overlap, and include an `other` option.
- **Use probabilities, not just labels.** `choice:probs` and `noul` let you set your own cut-offs and send low-confidence rows to a human.
- **Pin the model** (e.g. `jev-1.13.0`) before you tune thresholds, so that a new `jev-latest` doesn't move your cut-offs.
- **Mind cell limits.** An Excel cell holds at most 32,767 characters. For longer documents, split the text across a range, which is sent as a list.

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| `#NAME?` for `JEV.ASK` | The add-in isn't loaded in this Excel session. Open **Home ▸ Jev** (or reinstall) and recalculate. |
| Jev pane is blank, or Excel says the add-in couldn't start (local install) | The dev server isn't running. Run `npm run serve`, leave it running, and restart Excel. |
| No **Jev for Excel** under **SHARED FOLDER** (Windows) | The catalog must be a `\\server\share` network path with **Show in Menu** ticked, and Excel must be restarted after you add it. |
| `#VALUE!` "No Jev API key set" | Add your key in **Home ▸ Jev** and click **Save**. Saving recalculates the workbook. |
| `#VALUE!` "Invalid or missing API key (HTTP 401)" | Check your TypeSafe key in **Home ▸ Jev**. |
| `#N/A` "Could not reach the Jev API (network or CORS error)" | **Base URL** in **Home ▸ Jev** points at `api.typesafe.ai`, which browsers can't call. Clear it, click **Save** and try again. |
| `#N/A` "Could not reach https://… (network or CORS error)" | The server the add-in calls isn't answering. Locally, start it with `npm run serve`. When deployed, check the Worker with the `curl` command under [Proxy](#proxy). |
| `#N/A` "Rate limited" or "overloaded" | Jev was busy even after retries. Recalculate the affected cells a moment later. |
| Results didn't change after editing options | They should. Options are part of the cache key. If you changed the model or API key, click **Clear cache**. |
| `#BUSY!` for a long time | A very large fill-down is queued. Up to 6 requests run at once, and the task pane shows progress. |
| VBA: `#JEV! Could not reach https://…` | Check your network's proxy settings, or switch `HTTP_CLASS` to `"WinHttp.WinHttpRequest.5.1"`. |

---

## Development

```
jev-excel/
├── manifest.xml                 Add-in manifest (localhost URLs; the Worker serves it rewritten to its own URL)
├── wrangler.toml                Cloudflare Worker configuration
├── worker/index.mjs             Cloudflare Worker: serves dist/, proxies /v1/* to TypeSafe, rewrites manifest.xml and functions.json to its own URL
├── src/                         The add-in itself (built into dist/ by `npm run build` for the Worker)
│   ├── jev-core.js              Core logic: parsing, request building, batching, cache, retries (UMD, testable in Node)
│   ├── functions.js             Custom function bindings (JEV.ASK, CHOICE, SCORE, NOUL, STATE, MODELS)
│   ├── functions.json           Custom function metadata (names, parameters, help text)
│   ├── taskpane.html/.css/.js   Settings pane; shares a runtime with the functions
│   ├── cheatsheet.html          One-page formula guide (self-contained; served at /cheatsheet.html)
│   └── assets/                  icon.svg, and the icon PNGs rendered from it by `npm run icons`
├── vba/
│   ├── modJev.bas               VBA edition (=JEV, =JEV_STATE, =JEV_MODELS and macros)
│   └── ThisWorkbook.cls.txt     Optional Workbook_Open hook for function help text
├── scripts/
│   ├── serve.js                 HTTPS dev server on :3000, with the same /v1/* proxy as the Worker
│   ├── build.js                 Builds dist/ for the Worker
│   └── icons.js                 Renders src/assets/icon.svg to every PNG size
├── test/
│   ├── jev-core.test.js         Core logic against a mock Jev API
│   ├── proxies.test.js          Dev server and Worker: static files, proxying, URL rewriting
│   └── project.test.js          Version numbers agree; functions.json matches functions.js; help links land on the cheat sheet
└── .github/workflows/ci.yml     Runs the tests on Node 20 and 22
```

```bash
npm test            # all tests (Node 20+). No dependencies, network or API key needed.
npm run serve       # HTTPS dev server on :3000
npm run deploy      # build dist/ and deploy the Worker to Cloudflare
npm run icons       # re-render the icon PNGs after editing src/assets/icon.svg
npm run validate    # Microsoft's manifest validator (needs internet)
```

The add-in uses a **shared runtime**: `taskpane.html` loads `jev-core.js`, `functions.js` and `taskpane.js` into a single JavaScript context. That's why the settings pane and the formulas share one client, one cache and one set of statistics. `functions.json` is written by hand, so if you add a function, register it in both `functions.js` (`CustomFunctions.associate`) and `functions.json`. Give it a `helpUrl` pointing at a matching section of `cheatsheet.html`, which Marketplace review requires. `npm test` fails if any of these are missing or disagree. The same goes for the version number, which lives in `package.json`, `jev-core.js`, `manifest.xml` and `modJev.bas`.

**API reference used:** `POST /v1/systemone` with `{ model, state, questions }`. The question types are `noul` (optional `criteria.true`/`criteria.false`), `choice` (`criteria` maps each label to a description or `null`) and `score` (`criteria` is an array of 2–10 levels). Answers come back keyed by question id. `GET /v1/models` returns `{ models: [{ name, description, release_date }] }`. See the [TypeSafe API reference](https://docs.typesafe.ai/api) and the official [`@typesafe-ai/sdk`](https://www.npmjs.com/package/@typesafe-ai/sdk).

---

## Licence

[MIT](LICENSE) © 2026 Paul O'Brien. See also the [privacy policy](docs/PRIVACY.md), the [end-user licence agreement](docs/EULA.md) and [support](docs/SUPPORT.md). "Jev" and "TypeSafe" are names of TypeSafe AI and are used here only to describe compatibility.
