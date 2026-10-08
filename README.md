# Trip Split

A local-first bill splitter for trips and shared households. One HTML file,
no build step, served from GitHub Pages; your data lives in a Google Sheet you
own, behind a small Google Apps Script backend.

**Live app:** https://vlui93.github.io/trip-split/ — it opens on a Connect
screen, because it needs your own backend before it has anything to show.

## Features

- **Projects.** Keep a trip separate from everyday spending. A project with a
  date range opens automatically while you are inside those dates.
- **Splits.** Equal, exact amounts, percentages, or itemised line items with
  tax and tip shared in proportion to what each person ordered.
- **Currencies.** Any three-letter currency. Rates are fetched daily, and every
  expense keeps the rate it was saved with. You can override it per expense.
- **Settle up.** Works out the fewest transfers that clear every balance.
- **Recurring.** Expenses and transfers that repeat weekly, fortnightly,
  monthly, quarterly or yearly.
- **Receipt scanning and plain-English entry** (optional). Photograph a bill or
  describe it, and the form fills itself in for you to check.
- **Offline-first.** Every change is saved on the phone first and synced when a
  connection is available.
- Installs to the iOS home screen and runs full-screen.

## How it works

```
iPhone (index.html) ──HTTPS──▶ Apps Script web app ──▶ Google Sheet
      │                               │
  localStorage + write queue     optional: Gemini / Claude API
```

The app keeps a full copy of the data on the phone and never waits on the
network. Changes go into a queue that flushes to the backend when it is
reachable, then the phone pulls the latest state. Each phone works fully
offline, and the Sheet is the shared record.

## Setup

### 1. Create the Sheet and the script

Create the script **from inside the Sheet**. Starting at script.google.com
makes a standalone script with no Sheet attached (`useSheet(id)` can recover
one, see below).

1. Open <https://sheets.new> and give the spreadsheet a name.
2. Choose **Extensions → Apps Script**, delete the placeholder code, paste in
   all of [`Code.gs`](Code.gs) and save.
3. In the function dropdown, **select `setup` explicitly**, then click **Run**.
   The dropdown defaults to the first function in the file, so check it.
4. Approve the authorization prompt. Google shows *"Google hasn't verified this
   app"* for any personal script that makes web requests. Choose **Advanced →
   Go to (project) → Allow**.

`setup()` creates the tabs, a starting project and a starting currency (AUD),
and generates an API token. The token is printed in the execution log, and
`showToken()` prints it again. Re-running `setup()` is safe; it also migrates
older Sheets. If you update `Code.gs` and forget to re-run it, any tab or column
the new code needs is created on the next sync.

**Already made a standalone script?** Run `useSheet('<sheet id or URL>')` once,
then `setup()`.

### 2. Deploy it as a web app

1. **Deploy → New deployment**, click the gear next to *Select type*, choose
   **Web app**.
2. Set **Execute as: Me** and **Who has access: Anyone**.
3. Deploy and copy the URL ending in `/exec`.

For later updates, use **Deploy → Manage deployments → edit → Version: New
version**. A *new deployment* gets a new URL, and every phone would have to
reconnect.

### 3. Connect a phone

Open the live app in Safari, paste the `/exec` URL and the token, and tap
**Connect**. To install it: **Share → Add to Home Screen**.

Everyone in the group connects with the same URL and token. Each phone then
picks **More → Which one is you?**, which only sets the default payer and
highlights that person's row.

## Security model

**Why "Anyone" and not "Only myself".** A deployment restricted to your Google
account only answers requests that carry your Google session cookie. A page on
`github.io` is a different origin, so the browser never sends that cookie and
the request fails. Instead the deployment is reachable by URL, and the script
rejects every request without the shared token.

**Nothing secret is in this repository.** The backend URL and token are entered
on each device and kept in its `localStorage`. Anyone can load the page, but
without a token they cannot read or write anything.

**The token is the credential.** Treat it like a password. If it leaks, run
`resetToken()` and reconnect each phone. Everyone holding the token has full
access to that Sheet, which is the intended model for a small group that trusts
each other. It is not a permissions system.

Optional AI keys are stored in the script's properties, on Google's side. They
never reach the phone or this repository.

## Using it

### Projects

A project holds its own expenses, balances, settle-up, recurring items and
members. Tap the title on the home screen to switch projects, edit one or add a
new one. A project can have:

- **Dates or none.** On the first open of each day, the app switches to the
  project whose dates include today, or back to the ongoing project once a trip
  has ended. A project you pick by hand stays open for the rest of that day.
  Overlapping dates trigger a warning.
- **Currencies.** A project with only AUD hides all the conversion controls.
- **Places.** Free text, optional. Tapping one suggests its currency.
- **Members.** A fixed list of people, shown as a count in the header. A new
  trip starts with just you, so you tick who's going. A new ongoing project
  starts with everyone who exists at the time. **All** selects everyone
  currently on the roster.

Only a project's members are offered as payers and splitters on its expenses.
A person you add joins only the project you add them in: either the one whose
settings you are editing, or the current project when you add them under
**More → People**. A trip companion therefore never shows up in household
expenses, and a new housemate never shows up on a past trip.

Removing someone who already appears on a project's expenses shows a warning
first. They stay on those expenses, and their balance still shows.

Deleting a project deletes its expenses, settlements and recurring rules.

### Currencies and rates

Rates read as *1 AUD = 5.23 HKD*, the way people say them. They come from
[open.er-api.com](https://open.er-api.com), with
[Frankfurter](https://frankfurter.app) as the fallback. Fetched rates are cached
in the `ExchangeRates` tab for the day.

- **Add a currency** in **More → + Add currency**. The rate is fetched as you
  add it, so a mistyped code is rejected straight away.
- **Change the default** for one currency in Settings. It applies to new
  expenses until you clear it.
- **Override one expense** by typing over `1 AUD = …` on the add or edit
  screen. Only that expense uses the rate you type.
- **Editing keeps the original rate.** An expense opens at the rate it was saved
  with, so fixing a typo never re-converts it. Tap **Use today's** if you want
  it re-converted.

### Splitting

- **Equal.** Tap people out to exclude them.
- **Exact or percent.** Amounts are entered in the local currency, and the form
  shows how much is left to assign.
- **Itemised.** Line items are assigned to one or more people. Tax, tip and
  service go in one field and are shared in proportion to what each person
  ordered.

Shares are calculated in cents with largest-remainder rounding, so they always
add back to the total exactly.

### Settle up

The app works out the fewest transfers that clear every balance by matching
the largest debtor with the largest creditor, so a chain A→B→C collapses to
A→C. For three or fewer people this is provably minimal. **Mark as settled**
logs the payment.

### Recurring

- **A recurring expense** (rent, a subscription): on the add screen, set
  **Repeat** and optionally **Until**.
- **A recurring transfer** (a regular payment between two people): **Settle up
  → Set up a recurring transfer**.

A transfer settles a debt, so on its own it *creates* one. If a housemate pays
you toward rent each month, the rent needs to be a recurring expense too.

The home screen shows what is due next. From there you can change the amount or
end date, pause, resume or stop a series. Occurrences that have already been
added are ordinary expenses, and stopping a series keeps them.

### Receipt scanning (optional)

Two extra ways to fill in the add-expense form: **📷 Scan receipt** and
**💬 Describe it**. The second works with the iOS keyboard's dictation. Both
only fill in the form for you to check; nothing is saved until you tap Save.

They need an API key, set once from the script editor:

| | Set with | Cost | Trains on your data |
|---|---|---|---|
| Gemini (Google AI Studio) | `setGeminiKey('AIza…')` | free tier | yes, on the free tier |
| Claude (Anthropic) | `setApiKey('sk-ant-…')` | a few cents per receipt | no |

Anthropic API billing is separate from any Claude subscription. If both keys are
set, Claude is used. `testAi()` checks a key without needing a photo, and
`listGeminiModels()` lists the models your key can use.

Behind the scenes, if a Gemini model has been retired, the code switches to the
replacement named in Google's error, then to whatever the model list offers. If
the service is overloaded, it retries with backoff before moving to another
model. Photos are shrunk on the phone to 1568px before upload, sent once, and
never stored.

#### Offline receipt reading

When Google can't be reached (mainland China, no signal, a request hanging for
45 seconds) or no key is set, **📷 Scan receipt** reads the photo on the phone
with [Tesseract.js](https://github.com/naptha/tesseract.js) instead. Nothing is
uploaded. It picks out the total, line items, service charge and discounts with
plain pattern matching. Every item starts out shared by everyone, so you untick
whoever didn't have it. If the items don't add up to the total, it falls back to
an equal split and says so. The note field ("Sam didn't drink") is only used by
the AI.

Download it once while you have Wi-Fi: **More → Connection → Download offline
reader** (about 10 MB). The engine and English / Simplified / Traditional Chinese
data are kept in `ocr/` on this site, not a CDN, and `sw.js` caches them on the
phone so they load with no connection. The language is picked by currency: CNY
reads Simplified, HKD and MOP read Traditional. Expect it to misread some Chinese
characters; the numbers are what matters.

The parser has its own tests: `node tests/parse-receipt.test.mjs`.

## Design notes

**Offline and blocked networks.** The UI never waits on the network, so adding
expenses, splitting and settling up all work with no connection. Google services
are blocked in some countries, and there the queue simply holds until you are
back online or on a VPN. Fetching new rates and seeing other people's changes
both need a connection.

**Writes are idempotent.** Saving the same id twice updates the existing row
instead of adding another. Editing or deleting something another phone has
already deleted also succeeds. So a save that reached the Sheet but lost its
reply can be retried safely.

**One rejected change never blocks the rest.** If the backend turns down a
change, the phone keeps it, marks it, and carries on delivering everything
queued after it. The home screen shows what is waiting, with **Try again** and
**Discard**. Waiting changes stay visible on the phone after it syncs, because
they are layered back on top of the Sheet's copy. Going offline part-way through
a sync loses nothing.

**Dates use the spreadsheet's timezone.** Sheets stores a typed date at midnight
in the spreadsheet's timezone, and the Apps Script project has its own
timezone setting, which can differ. Dates are always read and written in the
spreadsheet's timezone, so they never shift by a day.

**Only the backend creates recurring occurrences.** It adds whatever is due on
every sync, under the script lock, using deterministic ids
(`rc_<rule>_<date>`). Each rule keeps a high-water mark that only moves forward.
As a result, several phones can never create the same occurrence twice, and a
deleted occurrence stays deleted.

Monthly rules keep their day. A series starting on the 31st falls on the last
day of shorter months and returns to the 31st after. A new occurrence appears on
the first sync on or after its date.

**Names are the key that links an expense to a person.** Renaming someone
rewrites every reference, including expenses, settlements, recurring templates
and project member lists. Removing someone who appears on existing expenses
shows a warning first.

## Sheet layout

| Tab | Columns |
|---|---|
| `Projects` | id, name, start_date, end_date, currencies, cities, members, archived |
| `People` | name |
| `Expenses` | id, date, city, description, paid_by, currency, amount_local, amount_aud, split_type, split_detail_json, item_breakdown_json, project_id |
| `ExchangeRates` | currency, rate_to_aud, last_updated, is_manual_override |
| `Settlements` | id, from, to, amount_aud, date, note, project_id |
| `Recurring` | id, project_id, kind, frequency, start_date, end_date, generated_through, paused, template_json |

`rate_to_aud` is AUD per one unit of the currency; the app shows the inverse.

## API

Every request is a `POST` with a `text/plain` body. That keeps it a CORS
"simple request", because Apps Script cannot answer a CORS preflight. The body
is `{ token, action, payload }`. Every response is either `{ ok: true, … }` or
`{ ok: false, error }`.

| Action | Payload |
|---|---|
| `bootstrap` | — (creates due recurring items, returns everything) |
| `addExpense` / `updateExpense` / `deleteExpense` | expense, or `{ id }` |
| `addSettlement` / `deleteSettlement` | settlement, or `{ id }` |
| `addProject` / `updateProject` / `deleteProject` | project, or `{ id }` |
| `addRecurring` / `updateRecurring` / `deleteRecurring` | rule, or `{ id }` |
| `setPeople` | `{ people: [names] }` |
| `setRate` / `clearOverride` / `refreshRates` | `{ currency, rate_to_aud }` / `{ currency }` / — |
| `addCurrency` / `removeCurrency` | `{ currency }` |
| `aiStatus` / `parseReceipt` / `parseText` | — / `{ image, … }` / `{ text, … }` |

The AI actions run outside the script lock, because they take seconds and touch
no rows.

## Troubleshooting

- **`setup()` "runs" but nothing happens.** The dropdown is set to a different
  function. Select `setup` and run it again.
- **"Backend returned non-JSON".** The deployment's access is not set to
  *Anyone*.
- **"The Sheet's script is older than this app."** The app needs a newer
  backend than the one deployed. Repaste `Code.gs`, then **Deploy → Manage
  deployments → edit → New version**, and tap **Try again** on the phone.
  Changes made in the meantime are kept on the phone.
- **"N changes couldn't sync."** The message under it gives the reason. **Try
  again** retries them; **Discard** drops them from that phone.
- **The home-screen icon doesn't change.** iOS caches it by URL, and deleting
  the home-screen item does not clear that cache. The icon filenames carry a
  version for this reason: bump `VERSION` in `mkicon.py`, run it, then reload
  the page in Safari before adding it again.

## Files

| File | |
|---|---|
| `index.html` | The whole frontend: markup, styles and logic |
| `Code.gs` | Apps Script backend |
| `manifest.webmanifest` | Web app manifest |
| `icon-*-v2.png`, `mkicon.py`, `icon.svg` | Home-screen icons and the script that generates them |
