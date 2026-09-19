# Trip Split

A mobile-first bill splitter. It started as one trip — 16 Oct – 1 Nov 2026 across
Hong Kong, Macau, Guangzhou, Chongqing and Chengdu — and now runs several
**projects** side by side: a trip with dates, and everyday AUD spending that
carries on around it.

Single-page vanilla HTML/CSS/JS, no build tooling. The frontend is served from
GitHub Pages; all data lives in a Google Sheet behind an Apps Script Web App.

- **Currencies:** AUD, HKD, MOP, CNY. Every expense is entered in local currency
  and converted to AUD for balance-keeping; both amounts stay visible everywhere.
- **Splits:** equal, exact amounts, percentages, and itemized (line items assigned
  to one or more people, with tax/tip prorated across whoever is on that receipt).
- **Settle Up:** computes the fewest transactions that clear every balance, so a
  chain like A owes B $40 and B owes C $40 collapses to A pays C $40.
- **Receipt photos and plain English:** optional — photograph a bill or describe
  it (dictation works) and it fills the form in for you to check. Runs on
  Gemini's free tier or on Claude. See §6.
- **Offline-first:** every change is written to the phone first and pushed to the
  Sheet when Google is reachable. This matters — see the mainland China note below.

---

## 1. Create the Sheet and the backend

> **Do not start at script.google.com.** That creates a *standalone* script with
> no Sheet attached. The script has to be created from inside the Sheet so it is
> bound to it. If you already made one there, see "Already started at
> script.google.com?" below — you don't have to throw it away.

1. Go to <https://sheets.new> and name the spreadsheet something like
   **Trip Split 2026**.
2. In that Sheet's menu bar, click **Extensions → Apps Script**. A new tab opens
   with a code editor and a file called `Code.gs` containing `function
   myFunction() {}`. Because you opened it from the Sheet, this script is bound
   to that Sheet — that is what makes `setup()` able to find it.
3. Select all of the placeholder code and delete it, paste in the whole of
   [`Code.gs`](Code.gs) from this repo, and save (⌘S).
4. In the function dropdown at the top of the editor, **explicitly select
   `setup`** — the dropdown defaults to whichever function is first in the file,
   so check it actually says `setup` before you click **Run**. Running the wrong
   function usually does nothing at all and reports success.

   Then click **Run**.
   - Google will ask you to authorize the script. This step has to be done by
     you, in the browser — it cannot be scripted. Choose your account, click
     **Advanced → Go to (project name) (unsafe)**, then **Allow**. The "unsafe"
     warning is just because the script is unpublished and yours alone.
   - `setup()` creates the four tabs with their headers, seeds three placeholder
     people and the four currency rows, generates an API token, and shows it.
5. Copy the API token from the dialog. If you clicked past it, everything
   `setup()` did is also printed in the **execution log** at the bottom of the
   editor — the Sheet name, the tabs created, the people, and the token. If you
   lose it later, run `showToken()` again.

**How to tell it worked:** the execution log ends with `API TOKEN: ...`, and your
Sheet has four tabs along the bottom. If the log just says "Execution completed"
with nothing above it, you ran a different function — check the dropdown.

### Already started at script.google.com?

A standalone script can still work — it just needs to be told which Sheet to use.

1. Open your Sheet and copy its id from the URL: the long string between `/d/`
   and `/edit`.
2. In the script editor, pick **`useSheet`** in the function dropdown. It needs
   an argument, which the dropdown cannot pass, so instead add this line
   temporarily at the bottom of the file:
   ```js
   function connectSheet() { useSheet('PASTE_THE_ID_HERE'); }
   ```
   Save, choose **`connectSheet`** in the dropdown, and **Run**.
3. Now pick **`setup`** and run that. From here the steps are identical.

`useSheet` also accepts the full Sheet URL if that is easier to paste. You can
delete `connectSheet` afterwards; the id is stored in script properties.

If you run `setup()` on a script that is attached to nothing, it stops with an
error telling you this rather than failing obscurely.

The tabs it creates:

| Tab | Columns |
|---|---|
| `Projects` | id, name, start_date, end_date, currencies, cities, members, archived |
| `People` | name |
| `Expenses` | id, date, city, description, paid_by, currency, amount_local, amount_aud, split_type, split_detail_json, item_breakdown_json, project_id |
| `ExchangeRates` | currency, rate_to_aud, last_updated, is_manual_override |
| `Settlements` | id, from, to, amount_aud, date, note, project_id |

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

Pages is already enabled and serving from `main` / root — nothing to turn on.
Pushes to `main` go live within a few seconds.

If it ever needs re-enabling (say the repo is recreated), do it once at
<https://github.com/vlui93/trip-split/settings/pages>: under **Build and
deployment → Source** choose **Deploy from a branch**, set branch **main** and
folder **/ (root)**, and **Save**. The first build takes a minute or two.

## 5. Add it to the iPhone home screen

1. Open the URL in **Safari** (this does not work from Chrome on iOS).
2. Tap the **Share** button (the square with the arrow).
3. Scroll down and tap **Add to Home Screen**.
4. Name it **Trip Split** and tap **Add**.

It launches full-screen with no Safari chrome, its own icon, and the dark theme
follows your phone's appearance setting.

**If the icon ever looks wrong**, the rule is that an `apple-touch-icon` must be a
**square, fully opaque, full-bleed** image with **no rounded corners of its own** —
iOS applies its own squircle mask. Rounding the artwork first leaves pale
triangles in the corners once the mask is applied. Transparency is just as bad,
since it composites onto black. `mkicon.py` writes RGB PNGs with no alpha channel
for exactly this reason.

**iOS caches home-screen icons by URL, and deleting the home-screen item does not
clear that cache.** Changing the bytes behind the same filename will not show up.
So the icon files carry a version in their name — bump `VERSION` in `mkicon.py`,
run it, and point `index.html` and the manifest at the new names. Then hard-reload
the page in Safari before adding it to the home screen again, or Safari may still
be reading the cached HTML with the old filenames in it. Your URL and token stay saved, so it
opens straight to the balances.

## 6. Optional: scan receipts and describe expenses out loud

Two extra ways to add an expense, both off by default:

- **📷 Scan receipt** — photograph the bill; the line items, total, tax and
  currency come back filled in.
- **💬 Describe it** — "Hotpot in Chengdu, 396 yuan, I paid, Mei didn't drink"
  becomes an itemized split. Tap the 🎤 on the iOS keyboard and say it instead of
  typing; that is ordinary system dictation, so it works in any text field.

Both land in the normal add-expense form for you to check. **Nothing is saved
until you tap Save**, and you can edit every field first.

### Turning them on — free option

They need an AI key. **Google AI Studio's free tier needs no credit card** and is
the default choice here.

1. Go to <https://aistudio.google.com/apikey>, sign in, **Create API key**. Copy it
   (it starts `AIza`).
2. In the Apps Script editor, add this temporarily at the bottom of `Code.gs`:
   ```js
   function connectAI() { setGeminiKey('AIza_PASTE_YOURS_HERE'); }
   ```
3. Select **`connectAI`** in the function dropdown, **Run**, then delete the
   function again. The key is now in script properties.
4. Run **`testAi`** from the same dropdown. It parses a sample sentence and prints
   the result to the execution log — so you can confirm the key works without
   photographing anything.
5. Reopen the app. The two buttons appear at the top of the add-expense sheet, and
   **More** shows `Gemini · free`.

Free-tier limits are per-minute and per-day and sit far above what a trip needs —
a few dozen receipts a day is not close. If you do hit one, the app says so and
you enter that expense by hand.

**Overload (503) is the common annoyance, not quota.** Free-tier traffic is the
first thing shed when Google is busy. The code retries three times with backoff
(honouring the `retryDelay` Google sends when it sends one), and if a model is
still overloaded after that it swaps to another model rather than failing —
a less fashionable one is usually free. Worst case it gives up after about ten
seconds and you type the expense in. If it happens constantly at the hour you
travel, Claude or Gemini's paid tier both skip the free-tier queue.

**Model names drift, and retired ones stay listed.** When a call fails on the
model, the code takes the replacement Google names in the error text first
(`Please update your code to use models/...`), then falls back to `ListModels`,
skipping anything it has already tried. It caches whatever works in the
`GEMINI_MODEL` script property.

If it ever runs out of candidates, the error names everything it tried. Run
`listGeminiModels()` in the script editor to print what your key can actually
use, then set `GEMINI_MODEL` directly in **Project Settings → Script
Properties**.

### The trade-off

Free tier means **Google may use what you send to improve their models**. For
restaurant receipts that is probably fine, but it is your call — the app says so
on the scan screen rather than burying it here. Avoid photographing anything with
a full card number. Google's paid tier drops the training use and still costs
well under a cent per receipt.

### Paid option — Claude

Better at messy receipts and mixed Chinese/English, and Anthropic does not train
on API traffic. Roughly **3–6 cents per receipt**, pay-as-you-go, **separate from
any Claude subscription** — Pro and Max do not include API credit.

1. Get a key at <https://console.anthropic.com> → API keys, add a little credit.
2. Same as above, but `setApiKey('sk-ant-...')`.

If both keys are set, Claude is used. `clearApiKey()` drops back to Gemini;
`clearGeminiKey()` turns everything off. The buttons disappear on their own when
no key is set.

### How it runs

| | Gemini | Claude |
|---|---|---|
| Model | `gemini-3.6-flash` (auto-corrected if retired) | `claude-opus-5` |
| Structured output | `responseSchema` (OpenAPI subset) | `output_config.format` JSON Schema |
| Cost | free tier | ~3–6c per receipt |
| Trains on your data | yes, on the free tier | no |

One schema is defined once and converted for Gemini — types upper-cased,
`additionalProperties` dropped, `propertyOrdering` added — rather than maintained
twice. Claude runs at `effort: "low"`; this is extraction, not hard reasoning, and
low effort keeps the round trip inside Apps Script's fetch timeout. Claude also
has server-side refusal fallbacks enabled, retrying once without that beta flag if
the API rejects it.

### What it won't do

Either key lives in the **script's** properties on Google's side — the phone never
sees it, and it is not in this repo.

These need a live connection, so they will not work in mainland China without a
VPN. Note that with Gemini both hops are Google, so one VPN covers it. Everything else in the app still works offline —
type the expense in by hand and it queues like any other change.

## 7. Putting it on all three phones

The Sheet is the shared store, so every phone runs the same setup:

1. Open **https://vlui93.github.io/trip-split/** in Safari.
2. Paste the same `/exec` URL and the same API token.
3. Add to Home Screen.
4. **More → Which one is you?** — pick your name.

That last step is per-phone. It defaults new expenses to you as the payer (which
is what you want nearly every time) and marks your row on the home screen. It is
stored on the device, not in the Sheet, so each phone can answer differently.

### How sharing behaves

**Syncing is not live.** A phone pulls when it opens, when you switch back to it,
when you pull down on the home screen, and when it flushes its own queue. So
Mei's dinner shows up on your phone next time you open it, not the instant she
saves it. For splitting a trip that is fine; do not expect a chat app.

**Adding is conflict-free.** Every expense gets a unique id on the phone that
created it, so three people adding at once never collide, even if all three were
offline and queued things up.

**Editing the same expense is last-write-wins.** If two of you edit the same
dinner before either syncs, the one that reaches the Sheet second overwrites the
other. Rare in practice, but it will not warn you.

**Everyone shares one token, so everyone has full access** — any phone can edit
or delete anything. That is the intended model for three people on a trip
together, not a permissions system. If someone loses their phone, run
`resetToken()` in the script editor and reconnect the other two.

**The roster is editable.** **More → People** adds and removes people, so a group
of two or five works as well as three. Renaming is safe: names are the key tying
an expense to a person, so a rename rewrites every reference — payer, split
participants, exact and percentage shares, itemised line items, logged
settlements, and this phone's "which one is you". Nothing is orphaned.

Removing someone who appears on existing expenses warns you first and tells you
how many. Their share stays on those expenses, so their balance keeps showing on
the home screen until you edit or delete them — which is the honest outcome, and
better than silently making their money disappear. Removing a person nobody has
shared an expense with is clean.

Since a roster change is shared, do it on one phone and let the others pull it
in. Two people renaming the same person at once is last-write-wins, like any
other edit.

## 8. Projects

A project is a set of expenses kept apart from the others, like a Splitwise
group. Balances, settle-up and history are always computed **within** one — trip
money and grocery money never mix.

Two exist out of the box:

| | Dates | Currencies | Places |
|---|---|---|---|
| **Greater China 2026** | 16 Oct – 1 Nov 2026 | AUD, HKD, MOP, CNY | the five cities + Sydney |
| **General** | ongoing | AUD | none |

Tap the title on the home screen to switch, edit, or add one.

### Dates do the switching for you

Give a project a date range and the app follows the calendar: while today falls
inside it, that project opens by default; outside it, you land back on the
ongoing one. So the trip picks itself up when you fly out and hands back to
everyday spending when you get home, with no switching to remember. You can
always override by hand.

Only one dated project can own a given day, so the editor warns you if the dates
you're setting overlap an existing project.

### What a project carries

- **Dates**, or ongoing.
- **Currencies.** A project with only AUD hides the conversion line and the rate
  tag entirely, which makes everyday entry two taps shorter.
- **Places.** Optional. With none, the place picker disappears and the history
  filter goes with it; with several you get the city chips. The receipt reader is
  told which places and currencies are valid for the project it's filling in.
- **Members.** Everyone by default; tap someone out to leave them off. Useful
  when the trip is three people but the flatshare is two.

Deleting a project deletes its expenses and settlements with it, and says so
first. Keeping at least one project is enforced.

### Upgrading an existing Sheet

`setup()` handles it and is safe to re-run. It adds the `Projects` tab, appends
`project_id` to `Expenses` and `Settlements` (appends, so nothing already in the
Sheet moves), widens a sheet that was trimmed to exactly its old column count,
seeds the two projects, and files every existing row into a project by its date —
rows inside a project's range go to it, everything else to the ongoing one.

**This one needs a redeploy** (Deploy → Manage deployments → pencil → New
version), because the backend changed.

---

## Exchange rates

Rates read the way you'd say them out loud: **1 AUD = 5.23 HKD**. That is how
they're shown everywhere and how you type an override.

Internally the Sheet stores `rate_to_aud` (AUD per one foreign unit), because
that is the number the conversion maths needs. The app flips it for display and
flips your input back on save, keeping eight decimal places so a rate you type
comes back exactly as you typed it.

On load the backend fetches the day's rates from
[open.er-api.com](https://open.er-api.com) — no API key — and caches them in the
`ExchangeRates` tab with a timestamp. If the cached rate is already from today it
is reused rather than re-fetched. [Frankfurter](https://frankfurter.app) is the
fallback for HKD and CNY; because its ECB feed does not carry MOP, MOP is then
derived from its de-facto 1.03 MOP : 1 HKD peg.

To match your bank's actual rate, go to **More**, tap a currency, and enter how
many of it one Australian dollar buys. It is flagged **manual** in amber everywhere instead of **live** in green,
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
| `icon-*-v2.png` | Home screen icons (180/167/152/120 for iOS, 192/512 for the manifest, 32 favicon). Versioned filenames — see below. |
| `mkicon.py` | Regenerates every icon PNG — bump `VERSION`, then `python3 mkicon.py` |
| `icon.svg` | Reference drawing of the icon |

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
| `addProject` / `updateProject` | project object | the saved project |
| `deleteProject` | `{ id }` | the id — also deletes its expenses and settlements |
| `addSettlement` | settlement object | the settlement |
| `deleteSettlement` | `{ id }` | the id |
| `aiStatus` | — | `{ ai: true/false, provider: 'gemini'\|'claude'\|'' }` |
| `parseReceipt` | `{ image, mediaType, city, currency, defaultPayer, hint }` | a draft expense |
| `parseText` | `{ text, city, currency, defaultPayer }` | a draft expense |

Every response is `{ ok: true, ... }` or `{ ok: false, error }`.

The three AI actions are handled before the script lock is taken — they touch no
rows and take tens of seconds, so holding the lock would block every other write
for the duration. Draft expenses are validated against the People list, the
currency list and the city list before they reach the app; a split the model
picked but did not populate falls back to an equal split rather than producing an
expense that splits to nothing.
