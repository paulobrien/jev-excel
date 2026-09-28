# Privacy policy

**Jev for Excel** · Effective 28 September 2026

This policy describes how the **Jev for Excel** service handles your information. Jev for Excel is an unofficial Excel add-in by Paul O'Brien ("I", "me") that lets you call TypeSafe AI's Jev model from worksheet formulas. It isn't affiliated with or endorsed by TypeSafe AI or Microsoft.

The policy applies to the whole Jev for Excel service:

- the Jev for Excel add-in for Excel on Windows, Mac and the web;
- the server that hosts the add-in and forwards its requests to TypeSafe (a Cloudflare Worker);
- the Jev for Excel VBA edition.

In short: **Jev for Excel doesn't collect, store, sell or share your personal information.** It sends your data only to TypeSafe, which you choose to use, and only to answer your formulas.

## Personal information

- **What personal information Jev for Excel handles:**
  - your TypeSafe API key;
  - any personal information contained in the text you choose to evaluate;
  - technical data, such as your IP address, that reaches the hosting server as part of any internet request.

  It doesn't ask for your name, email address or any other account details.
- **How it's used:** only to answer your formulas. Your text and key go to TypeSafe for that purpose and nothing else. Nothing is used for analytics, advertising, profiling or training.
- **Retention:** Jev for Excel keeps none of it on its servers. Your key and settings stay on your own device until you delete them. Answers are held in memory only until Excel closes.
- **Sharing:** your information isn't sold, rented or shared with anyone. It goes only to the service providers named under [Third parties](#third-parties), and only as needed to run the service.
- **Your rights:** Jev for Excel holds no personal information about you, so there's nothing to access, correct or erase on its side. [Your control](#your-control) explains how to delete what's on your device. To exercise your rights over data held by TypeSafe, Cloudflare or Microsoft, contact them directly.

## What Jev for Excel handles, and where it goes

| Data | What happens to it |
|---|---|
| **Text you ask about.** This is the `text` argument of a `JEV` formula, or the fields built by `JEV.STATE`, together with your options and question. | Sent to TypeSafe's API (`api.typesafe.ai`) to get an answer. The hosted add-in sends it through its Cloudflare Worker, which forwards it immediately and doesn't keep it. TypeSafe processes it under [its own terms and privacy policy](https://typesafe.ai). |
| **Your TypeSafe API key** | Stored only on your device, in the add-in's browser storage (`localStorage`). The VBA edition stores it in your Windows user registry. It's sent with each request so TypeSafe can authenticate you, and the Worker passes it through without keeping it. It is never written into your workbooks. |
| **Answers from Jev** | Returned to your worksheet. They're also cached in memory so recalculating doesn't call TypeSafe again. The cache is cleared when Excel closes. |
| **Settings** (model name, base URL, cache on/off) | Stored only on your device, next to your key. |

## What isn't collected

- **No analytics, telemetry, tracking or advertising.** The add-in sets no cookies.
- **Nothing is logged or stored by the Jev for Excel server's code.** It only forwards requests to TypeSafe and serves the add-in's files.
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

If the way Jev for Excel handles your information changes, this policy will be updated at the same address with a new effective date.

## Contact

For questions about this policy, see [Support](SUPPORT.md).
