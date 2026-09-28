# Privacy policy

**Jev for Excel** · Effective 28 September 2026

Jev for Excel is an unofficial, open-source Excel add-in by Paul O'Brien ("I", "me"). It lets you call TypeSafe AI's Jev model from worksheet formulas. It isn't affiliated with or endorsed by TypeSafe AI or Microsoft.

This policy covers the add-in, the Cloudflare Worker that hosts it, and the VBA edition in this repository. In short: **I don't collect, store, sell or share your data.** The add-in sends your data only to TypeSafe, which you choose to use, and only to answer your formulas.

## What the add-in handles, and where it goes

| Data | What happens to it |
|---|---|
| **Text you ask about.** This is the `text` argument of a `JEV` formula, or the fields built by `JEV.STATE`, together with your options and question. | Sent to TypeSafe's API (`api.typesafe.ai`) to get an answer. The hosted add-in sends it through its Cloudflare Worker, which forwards it immediately and doesn't keep it. TypeSafe processes it under [its own terms and privacy policy](https://typesafe.ai). |
| **Your TypeSafe API key** | Stored only on your device, in the add-in's browser storage (`localStorage`). The VBA edition stores it in your Windows user registry. It's sent with each request so TypeSafe can authenticate you, and the Worker passes it through without keeping it. It is never written into your workbooks. |
| **Answers from Jev** | Returned to your worksheet. They're also cached in memory so recalculating doesn't call TypeSafe again. The cache is cleared when Excel closes. |
| **Settings** (model name, base URL, cache on/off) | Stored only on your device, next to your key. |

## What isn't collected

- **No analytics, telemetry, tracking or advertising.** The add-in sets no cookies.
- **Nothing is logged or stored by the Worker's code.** It only forwards requests to TypeSafe and serves the add-in's files.
- **Nothing is read from your workbook** except the cells you pass to a `JEV` formula.
- **No account with me.** You don't sign up for anything, and I don't know who uses the add-in.

## Third parties

These services process data when you use the add-in, each under its own policy:

- **TypeSafe AI** receives the text, questions and API key described above, and bills your TypeSafe account for usage. See [typesafe.ai](https://typesafe.ai).
- **Cloudflare** hosts the add-in and the Worker. As the network provider, it handles every request and may process technical data such as IP addresses and request metadata. See the [Cloudflare privacy policy](https://www.cloudflare.com/privacypolicy/).
- **Microsoft** provides Excel and the Office JavaScript library (`office.js`), which the add-in loads from Microsoft's servers. See the [Microsoft Privacy Statement](https://privacy.microsoft.com/privacystatement).

## Your control

- To stop sending data, remove the add-in from Excel, or delete the `JEV` formulas.
- To delete your key and settings from the add-in, clear the fields in **Home ▸ Jev** and click **Save**. Removing the add-in also removes them. In the VBA edition, run `JevSetApiKey` with a blank key.
- For data held by TypeSafe, contact TypeSafe directly.
- Don't send data that your organisation's policies don't allow you to share with a third-party API.

## Children

The add-in isn't aimed at children, and I don't knowingly handle children's data.

## Changes

If this policy changes, the new version will be published here with a new effective date. The repository's history shows every earlier version.

## Contact

For questions about this policy, see [Support](SUPPORT.md).
