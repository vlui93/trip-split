# Trip Split — Greater China 2026

A mobile-first bill splitter for three people, built for the 16 Oct – 1 Nov 2026 trip
(Hong Kong · Macau · Guangzhou · Chongqing · Chengdu, home base Sydney).

Single-page vanilla HTML/CSS/JS, no build tooling. The frontend is served from
GitHub Pages; all data lives in a Google Sheet behind an Apps Script Web App.

- **Currencies:** AUD, HKD, MOP, CNY. Every expense is entered in local currency
  and converted to AUD for balance-keeping; both amounts stay visible everywhere.
- **Splits:** equal, exact amounts, percentages, and itemized (line items assigned
  to one or more people, with tax/tip prorated across whoever is on that receipt).
- **Settle Up:** computes the fewest transactions that clear every balance, so a
  chain like A owes B $40 and B owes C $40 collapses to A pays C $40.
- **Offline-first:** every change is written to the phone first and pushed to the
  Sheet when Google is reachable. This matters — see the mainland China note below.

---

## 1. Create the Sheet and the backend

1. Go to <https://sheets.new> and name the spreadsheet something like
   **Trip Split 2026**.
2. **Extensions → Apps Script**. This creates a script bound to that Sheet.
3. Delete the placeholder `myFunction` code, paste in the whole of
   [`Code.gs`](Code.gs), and save (⌘S).
4. In the function dropdown at the top, pick **`setup`** and click **Run**.
   - Google will ask you to authorize the script. This step has to be done by
     you, in the browser — it cannot be scripted. Choose your account, click
     **Advanced → Go to (project name) (unsafe)**, then **Allow**. The "unsafe"
     warning is just because the script is unpublished and yours alone.
   - `setup()` creates the four tabs with their headers, seeds three placeholder
     people and the four currency rows, generates an API token, and shows it.
5. Copy the API token from the dialog (or from **View → Logs**). If you lose it,
   run `showToken()` again.

The tabs it creates:

| Tab | Columns |
|---|---|
| `People` | name |
| `Expenses` | id, date, city, description, paid_by, currency, amount_local, amount_aud, split_type, split_detail_json, item_breakdown_json |
| `ExchangeRates` | currency, rate_to_aud, last_updated, is_manual_override |
| `Settlements` | id, from, to, amount_aud, date, note |

## 2. Deploy it as a Web App

1. **Deploy → New deployment**. Click the gear next to "Select type" and choose
   **Web app**.
2. Set:
   - **Execute as:** Me
   - **Who has access:** **Anyone**
3. **Deploy**, authorize again if prompted, and copy the **Web app URL**. It ends
   in `/exec`.

### Why "Anyone" and not "Only myself"

The brief asked for "accessible to me only", and that is the right instinct, but
it cannot work here. A deployment restricted to your account only answers
requests that carry your Google session cookie. A page served from
`github.io` is a different origin, so the browser sends no Google cookie and
Apps Script replies with a redirect to the sign-in page with no CORS headers.
The fetch fails before it ever reaches your script. There is no arrangement of a
static site plus a private Apps Script deployment that avoids this.

So the deployment is reachable by URL, and the script instead checks a **shared
token** on every single request and rejects anything else. The token is the
credential; the URL alone gets you an error.

That makes the token worth protecting, which is why **neither the URL nor the
token is committed to this repo**. The app asks for both on first load and keeps
them in `localStorage` on your phone. A public repo therefore leaks nothing — it
is just the frontend code. Anyone can load the page, but with no token they
cannot read or write a single row.

Two things follow from that:

- If you ever paste the token into a chat, an issue, or a screenshot, run
  `resetToken()` in the script editor and reconnect each phone.
- Each of the three of you enters the same URL and token once. That is the
  intended sharing model.

If you would rather not have a bearer token in three people's browsers at all,
the alternative is deploying "Anyone with a Google account" and having everyone
sign in — but that reintroduces the cross-origin cookie problem for a static
page, so it would mean hosting the frontend inside Apps Script itself with
`HtmlService` and dropping GitHub Pages. That is a real option; it is just a
different app, and it loses the home-screen polish and the instant loads.

**Redeploying:** after editing `Code.gs`, use **Deploy → Manage deployments →
edit (pencil) → Version: New version → Deploy**. Creating a *new* deployment
instead gives you a new URL and you would have to reconnect every phone.

## 3. Point the app at it

Open the Pages URL (below) on your phone. The first screen asks for the
**`/exec` URL** and the **API token**. Paste both and tap **Connect**. It will
pull the Sheet, fetch today's rates, and land on the home screen.

Then go to **More → People** and replace the three placeholder names with the
real ones. Do this before adding expenses — splits are stored against names.

## 4. Live URL

**https://vlui93.github.io/trip-split/**

If GitHub Pages is not already on, turn it on once:

1. Open <https://github.com/vlui93/trip-split/settings/pages>
2. Under **Build and deployment → Source**, choose **Deploy from a branch**.
3. Branch: **main**, folder: **/ (root)**. **Save**.
4. Wait ~1 minute, then reload the URL above. The first build can take a couple
   of minutes; after that, pushes go live in seconds.

## 5. Add it to the iPhone home screen

1. Open the URL in **Safari** (this does not work from Chrome on iOS).
2. Tap the **Share** button (the square with the arrow).
3. Scroll down and tap **Add to Home Screen**.
4. Name it **Trip Split** and tap **Add**.

It launches full-screen with no Safari chrome, its own icon, and the dark theme
follows your phone's appearance setting. Your URL and token stay saved, so it
opens straight to the balances.

---

## Exchange rates

Rates are **AUD per one unit** of the currency, so 1 HKD ≈ 0.19 AUD.

On load the backend fetches the day's rates from
[open.er-api.com](https://open.er-api.com) — no API key — and caches them in the
`ExchangeRates` tab with a timestamp. If the cached rate is already from today it
is reused rather than re-fetched. [Frankfurter](https://frankfurter.app) is the
fallback for HKD and CNY; because its ECB feed does not carry MOP, MOP is then
derived from its de-facto 1.03 MOP : 1 HKD peg.

To match your bank's actual rate, go to **More**, tap a currency, and enter your
own. It is flagged **manual** in amber everywhere instead of **live** in green,
and the daily refresh will not touch it until you tap **Clear override**.

**Expenses keep the rate they were saved with.** Changing a rate today does not
silently rewrite what last week's dinner cost. The rate used is shown on each
expense's detail view.

## How the splits work

- **Equal** — tap people off to exclude them.
- **Exact** — type each person's share in the local currency. It has to add up to
  the total, and the form tells you how much is left to assign.
- **Percent** — same, adding to 100%.
- **Itemized** — add line items and tap which people are on each one. Shared
  items are divided among whoever is tapped. Tax, tip and service go in one
  field and are prorated by what each person ordered, so the person who had the
  ¥240 hotpot carries more of the service charge than the person who had a ¥48
  beer. The receipt total is computed from the items, so it cannot drift.

Shares are computed in AUD cents with largest-remainder rounding, so the parts
always add back to the expense total exactly — no lost or invented cents.

## Mainland China

`script.google.com` is blocked in mainland China. Guangzhou, Chongqing and
Chengdu are the bulk of this trip, so plan for the backend to be unreachable for
most of it.

The app is built for that. Everything is stored on the phone first and the UI
never waits on the network, so adding expenses, splitting, and the settle-up
view all work fully offline. Pending writes queue up and the home screen shows
how many are waiting. They flush automatically the moment Google is reachable
again — over a VPN, or back in Hong Kong or Sydney.

What will *not* work offline: fetching new exchange rates, and seeing the other
two people's expenses. If all three of you are entering expenses in Chengdu,
nobody sees anyone else's until someone's queue syncs. Set the rates before you
cross the border, or set them manually, and one person on a VPN syncing daily
keeps everyone current.

## Files

| File | What |
|---|---|
| `index.html` | The whole frontend — markup, styles, and logic in one file |
| `Code.gs` | Apps Script backend; paste into the Sheet-bound script |
| `manifest.webmanifest` | Web app manifest for the home-screen install |
| `icon-180.png` | Home screen icon |
| `icon.svg` | Source for the icon |

## API

All requests are `POST` with a `text/plain` body (which keeps them CORS-simple —
Apps Script cannot answer a preflight) shaped as
`{ token, action, payload }`.

| Action | Payload | Returns |
|---|---|---|
| `bootstrap` | — | people, expenses, rates, settlements |
| `addExpense` / `updateExpense` | expense object | the saved expense |
| `deleteExpense` | `{ id }` | the id |
| `setPeople` | `{ people: [names] }` | the names |
| `setRate` | `{ currency, rate_to_aud }` | all rates |
| `clearOverride` | `{ currency }` | all rates |
| `refreshRates` | — | all rates |
| `addSettlement` | settlement object | the settlement |
| `deleteSettlement` | `{ id }` | the id |

Every response is `{ ok: true, ... }` or `{ ok: false, error }`.
