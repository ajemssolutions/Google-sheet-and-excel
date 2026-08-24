# AJEMS Google Sheets Connector — localhost prototype

Proves the whole chain works against real AJEMS data: Google login → pick a
Sheet → connect it to an AJEMS form → map columns → sync → keep syncing on a
schedule, including updates to rows that were already sent.

Multi-task from the start. One task = one sheet tab → one AJEMS form. Add as
many as you like; each has its own mapping, schedule and history.

---

## 1. Run it

```
npm install
cp .env.example .env      # then fill in the three Google values (step 2)
npm start
```

The Google setup is only needed for Google Sheets. To try the connector with an
Excel file, connect AJEMS and upload one - nothing else is required.

Open **http://localhost:3000**.

Requires Node 20 or newer. The only dependency is Express — everything else
uses built-in `fetch`.

---

## 2. Google Cloud setup (about 10 minutes, one time)

1. Go to **console.cloud.google.com** → create a project (or pick an existing one).
2. **APIs & Services → Library** → enable these three:
   - Google Sheets API
   - Google Drive API
   - Google Picker API
3. **APIs & Services → OAuth consent screen**
   - User type: **External**
   - Fill in app name, your email, developer email
   - **Add yourself as a Test user.** While the app is unverified only test
     users can sign in — that is fine for a prototype and needs no review.
4. **APIs & Services → Credentials → Create credentials → OAuth client ID**
   - Type: **Web application**
   - Authorised redirect URI: `http://localhost:3000/oauth/callback`
     (must match exactly, including the port)
   - Copy the client ID and secret into `.env`
5. **Create credentials → API key** — copy it into `GOOGLE_API_KEY` in `.env`.
   Google's file picker will not open without one. This is separate from the
   OAuth client, and it is developer configuration: an end user should never be
   asked for it.
Restart the server after editing `.env` — it is only read at startup.

Configuration problems are printed to the terminal at startup, not shown in the
app - a user has no use for the name of an environment variable. If Google is
not configured, the app simply says Google Sheets are unavailable and carries
on with Excel uploads.

### "Error 401: invalid_client"

Google is saying it does not recognise the client ID being sent. It is always a
credentials problem, never an app problem. In order of likelihood:

1. **The server was not restarted** after `.env` was edited.
2. **The API key was pasted instead of the OAuth client ID.** The client ID ends
   with `.apps.googleusercontent.com`; an API key starts with `AIza`.
3. **The OAuth client was deleted**, or lives in a different Google Cloud
   project than the one where you enabled the APIs.
4. **A stray space, quote or line break** got copied in with the value.

The app now checks all of this before sending you to Google and names the
offending value, so you should see a readable page rather than Google's.

---

## 2b. Running it on a server

Everything above assumes localhost. On a real domain four things change, and
missing any one of them shows as **"Google sign-in is not set up on this
server yet"**.

**1. `.env` is not in the repository.** It is git-ignored on purpose, so the
secret never reaches GitHub. Create it on the server:

```
GOOGLE_CLIENT_ID=...
GOOGLE_CLIENT_SECRET=...
GOOGLE_API_KEY=...
GOOGLE_REDIRECT_URI=https://your-domain/oauth/callback
PORT=3000
```

Restart after writing it - the file is only read at startup, and the terminal
prints exactly what is missing.

**2. Add the live redirect URI in Google Cloud.** Credentials > your OAuth
client > Authorised redirect URIs > add `https://your-domain/oauth/callback`.
It must match `GOOGLE_REDIRECT_URI` character for character. Google only
accepts https on a public host.

**3. Loosen the API key restriction.** If the key is restricted to
`http://localhost:3000/*` the file picker will not open on the live domain.
Credentials > your API key > Website restrictions > add `https://your-domain/*`.

**4. Publish the OAuth app**, or only listed test users can sign in. Google
Auth Platform > Audience > Publish app. Because the app only asks for
`drive.file`, this needs basic verification, not a security assessment.

The AJEMS side needs none of this. The workspace URL and secret key are typed
into the UI and stored on the server, so an Excel-only deployment works with no
Google setup at all.

## 3. The scope this uses

**`drive.file` and nothing else.**

The user picks a spreadsheet in Google's own file dialog. That act of picking
is what grants this app access to that one file. Nothing else in their Drive is
readable, and there is no other way in.

Why it matters commercially: Google classes `drive.file` as **non-sensitive**,
so publishing needs only basic app verification — no annual CASA security
assessment. `drive.readonly`, which would allow listing every spreadsheet in an
account, is **restricted**: publishing that costs roughly $500–4,500 a year,
every year, plus weeks of review.

`spreadsheets.readonly` is deliberately absent too. `drive.file` is itself a
valid Sheets API scope, so the Sheets API reads a picked file's values under
it. Adding `spreadsheets.readonly` would pull a *sensitive* scope back in for
no capability the app does not already have.

**The picker must be told the Cloud project number.** `PickerBuilder.setAppId`
is required for `drive.file` — without it the pick looks like it worked, the
app is told the file id, and then every read fails with *Requested entity was
not found*, because the file was never actually shared with the app. The
project number is read from the start of `GOOGLE_CLIENT_ID`, so there is
nothing extra to configure; `GOOGLE_PROJECT_NUMBER` in `.env` overrides it if
the client ID is ever an unusual shape.

**Only Google Sheets are selectable.** An `.xlsx` uploaded to Drive is a
different file type that the Sheets API cannot read, so the picker filters it
out and the callback rejects it with a clear message rather than failing later.

**Grants accumulate.** Pick one sheet, then another; both stay connected and
nothing has to be disconnected. That is what makes multi-task work, and it is
why a grant survives after the picker closes — the scheduler keeps reading the
file long afterwards.

The one thing this scope cannot do is watch a folder for *newly created*
sheets, since a per-file grant cannot cover a file that does not exist yet.

### "Error 403: access_denied" for a second account

While the app is in **Testing**, every account that signs in must be listed
individually under Google Auth Platform → Audience → Test users. Adding one
account does not cover the rest — a different Gmail address is a different
test user. Add it there, or publish the app.

### Going live for any Google user

| Publishing status | Who can sign in | Cost |
|---|---|---|
| Testing (default) | Up to 100 listed test users; **consent expires after 7 days** | free |
| Internal | Only accounts in your Workspace org; no verification ever | free |
| **Production + `drive.file`** | **Anyone with a Google account** | basic verification — free |
| Production + `drive.readonly` | Anyone | verification **and** annual CASA |

Basic verification wants an app name, logo, a homepage and privacy policy on a
domain you own, and domain ownership proven in Search Console. It is paperwork,
not code, so it can start in parallel.

## 4. The flow

**Connections** — AJEMS first (tenant name + secret key), then section 2,
**Google Sheet / Excel sheet**, which is where a source comes from:

- **Continue with Google** - connects your account so Google Sheets can be
  picked and synced on a schedule.
- **Upload Excel sheet** - takes an `.xlsx`, `.xlsm`, `.xls` or `.csv` straight
  from this computer. No Google account needed. Uploaded files are listed
  underneath and shared by every task, so one file can feed several; a file
  cannot be removed while a task still uses it.

Uploading a file opens the task flow straight away with that file already
chosen - the same five steps a Google Sheet goes through. Each file in the list
also has its own **New task** button.

Step 1 of the wizard starts with the same choice: **Google Sheet** or **Excel or
CSV file**. Choosing Google without a signed-in account shows a short message
and a **Continue with Google** button right there, rather than refusing to open
the wizard. Choosing Excel offers **Upload Excel sheet** plus a dropdown of
files uploaded before.

An uploaded file has no schedule - nothing on disk changes between uploads, so
the task imports once when saved, and **Sync now** repeats it. Google Sheets
keep their schedule as before.

**Connections, continued** —  One call
to `workspace_config` both tests the key and loads every app and form. Three
failures are told apart: a wrong key (the API answers with JSON, so it stops
instead of trying more hosts), a host that is not the API at all (HTML or a
proxy error, so it keeps trying), and nothing reachable (the error lists every
host it tried and why).

**Tasks → Add task** — five steps:

1. **Sheet** — name the task, then **Choose from Google Drive**. Google's own
   dialog opens with My Drive, Shared with me and Recent, filtered to Google
   Sheets. Picking a file is what grants the app access to it. Then tick one or
   more **tabs**. Each tab is read the moment you tick it, showing the detected
   AJEMS field type per column, how many cells are filled, and sample values,
   right there on the step. Row 1 is always the header.
2. **App** — pick an application from the grid, or create one (name,
   description, icon, colour). No Next button: choosing is what moves you on.
3. **Forms** — **one form per tab.** Each tab gets its own row: pick an
   existing form from the dropdown, or press *+ New*, which opens a name box
   pre-filled with the tab name (Enter creates it). **Create a form for every
   tab** does all the outstanding ones at once, naming each after its tab.
4. **Mapping** — always shown, never skipped. Columns are matched to fields by
   name (case and punctuation ignored, so `Mobile No.` finds `mobile no`), and
   every one of them can be changed. With several tabs there is a bar along the
   top to switch between them, showing how many columns each has mapped.
5. **Schedule** — how often it runs, and how a row is recognised.

**Saving a task syncs it straight away**, then it keeps running on its schedule.
Each scheduled check is one small Drive call for `modifiedTime`; the sheet is
only read when it actually changed. **Sync now** on the task card forces a run
without waiting. Watch it all on the **Activity** page.

**Edit** opens straight at the schedule — the sheet, form and mapping are
already settled, so there is nothing to re-walk.

**Updates** — the point of the prototype. Let a task sync, edit a cell in Google
Sheets, wait for the next run. The record is **updated**, not duplicated. Add a
row and only that row goes up.

## 4b. Tasks belong to an organisation

A task is saved against the AJEMS tenant that created it. Connect a different
workspace and you see that workspace's tasks; connect the first one again and
its tasks come back. Scheduled runs are scoped the same way, because the secret
key in hand only works for its own tenant - another organisation's tasks would
fail if they ran.

## 4c. Re-importing an Excel task

An uploaded file cannot change on its own, so an Excel task carries **Upload
new file & sync** on its card. Pick a newer copy of the same spreadsheet and it
uploads, replaces the file behind the task, and syncs in one go.

Nothing is duplicated. The dedupe key comes from the tab name and the row's own
values, never from the file, so the task's history survives the swap:

| What is in the new file | What happens |
|---|---|
| A row that was already sent | Left alone |
| A row that is new | Created in AJEMS |
| A row whose values changed | Updated in place, if identity is a key column |

If the new file is missing a sheet the task maps, it is refused by name and the
old file stays in place, rather than syncing nothing and looking like it worked.

**Sync again** re-runs against the current file without uploading.

## 4d. Setup guides

The sidebar carries a walkthrough for each source, **Google Sheets** and
**Excel & CSV**, built from the same components as the Tally Connector setup
page so the connectors read as one product: an integration header, a "before
you begin" list, four numbered steps each with a screenshot beside it, and a
closing call to action.

Both guides come from one data structure in `public/app.js` (`GUIDES`) and one
renderer, so changing a step is editing a string. Screenshots live in
`public/shots/` and a step lists its own:

```javascript
shots: [['Field mapping', 'mapping.png']]
```

A step can carry two, which stack in the media column. The Sheets guide runs to
seven steps and the Excel guide to six; several screenshots are shared between
them, since the app, forms and mapping steps are identical whichever source the
rows came from.

## 5. Row identity — the thing to form an opinion on

AJEMS has no upsert, and a sheet row number is not a stable identifier: insert
one row at the top and every row below shifts down. So the connector derives a
key from the row's own content, and keeps a `key → AJEMS response id` table.

**Key column(s)** — you nominate the column that identifies a row (invoice
number, email, SKU). Editing any other cell is recognised as an update to the
same record. This is the mode to use for sheets people edit.

**Whole row** — the key is a hash of every mapped value. No setup, but an
edited row looks like a brand-new row and gets inserted again. Fine for
append-only sheets, wrong for anything else.

The tab is always part of the key, whichever mode you choose, because each tab
targets its own form: the same values in two tabs are two records in two
different places.

Values are normalised before hashing, so `1,000` and `1000`, and differences
in case or spacing, do not produce false "changed" results.

"Sync now" seeds the table from what is already in the AJEMS form first, so a
restart or a second machine does not re-send everything.

Test 9 in the suite covers the case that matters: reverse the row order so
every row number changes, sync, and nothing is duplicated.

---

## 6. Automated tests

```
npm test
```

**`test/run-test.js`** — 45 checks on the libraries directly: detection, URL
normalisation, key rejection, first sync, re-run with no changes, an edited row
updating in place, an appended row, reordered rows, both identity modes,
seeding from AJEMS, per-row error reporting.

**`test/run-http-test.js`** — 35 checks through the actual HTTP routes the
browser calls. This suite exists because the first one passed while the UI was
broken: the mock used `id` where the real API uses `app_id`, so the application
dropdown had empty values and the wizard said "choose an application first"
with an application visibly selected. Anything the UI depends on is now checked
end to end, and the mock's field names match the API exactly. It also asserts
that the OAuth URL requests `drive.file` and never `drive.readonly` or a
`spreadsheets` scope, so a restricted scope cannot creep back in unnoticed.

Both run against a stand-in AJEMS server that also reproduces the real nginx
behaviour where a URL without a trailing slash answers 301. No Google
credentials needed; the Sheets reader is stubbed.

---

## 7. What this prototype is not

Deliberately out of scope, and each is a real gap before this becomes a
product:

- **Single user.** One Google account and one AJEMS workspace, in one
  `data.json`. No multi-tenancy, no per-user isolation.
- **Tokens are stored in plain text** in `data.json`. Fine on your own
  machine, not acceptable on a server — production needs encryption at rest
  and a key-management story.
- **No token-revocation handling.** If the Google grant is revoked, tasks fail
  and auto-pause after three consecutive failures, but nobody is notified.
- **Deleted sheet rows are ignored.** The AJEMS record stays. That is the safe
  default, but it is a decision to confirm, not an oversight.
- **The scheduler is an in-process timer.** It stops when the server stops, so
  leave the terminal open. A hosted version needs a real job queue.
- **No manual sync.** Tasks run only on their schedule; the shortest is five
  minutes, so testing a change means waiting for the next run.
- **No concurrency control.** Two syncs of the same task cannot overlap, but
  nothing stops AJEMS being edited underneath a sync.
- **`field_type` vocabulary is partly guessed.** Confirmed values are Text,
  Phone, Number, Date and Dropdown; anything uncertain falls back to Text,
  which always works. Worth getting the full list from the AJEMS team, since a
  wrong value fails silently.

---

## 8. Layout

```
server.js            Express app, all routes
lib/env.js           tiny .env reader (no dotenv dependency)
lib/store.js         JSON persistence: settings, tasks, link table, log
lib/google.js        OAuth, token refresh, Drive and Sheets calls
lib/ajems.js         JSON Builder client — trailing slashes, no redirect following
lib/detect.js        column → AJEMS field_type detection
lib/engine.js        row identity, create vs update, dedupe, concurrency pool
lib/scheduler.js     cheap change detection, staggered polling, auto-pause
public/setup.css     setup-guide components, shared look with Tally Connector
public/shots/        guide screenshots, cropped to the panel they show
                     (namespaced .gstep so it cannot collide with the wizard)
public/              single-page UI
test/                mock AJEMS server + 35 checks
```

`lib/engine.js` knows nothing about Google Sheets — it takes rows and a
mapping. That is the same shape as the Excel connector's engine, so whichever
direction this goes (in-AJEMS, extension, Apps Script), the sync logic moves
with it.
