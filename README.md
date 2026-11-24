# Jobo for Sheets

Google Sheets add-on that imports filtered Jobo job data, and the hiring signals of a list of companies, into
the active spreadsheet. Two menu items under **Extensions**: **Import jobs…** and **Import hiring signals…**.

Full setup and filter docs: [jobo.world/docs/connectors/google-sheets](https://jobo.world/docs/connectors/google-sheets).

Of the marketplaces surveyed, Google Workspace Marketplace is the only high-authority listing that emits a
**dofollow** link to the vendor site — which is why this exists alongside the connectors that were picked
purely for distribution.

## Scopes — the cost driver

```
https://www.googleapis.com/auth/spreadsheets.currentonly
https://www.googleapis.com/auth/script.container.ui
https://www.googleapis.com/auth/script.external_request
```

`spreadsheets.currentonly` grants access to the open document only, never the user's Drive.
`script.external_request` is unavoidable — it is what permits any outbound call, so without it the add-on
cannot reach the API at all. The plan originally assumed two scopes; this is the third.

> ⚠️ **Confirm the scope classification before submitting.** Restricted scopes trigger a CASA Tier 2
> third-party security assessment, which is slow and expensive; sensitive scopes do not. These three are
> believed to fall outside the restricted families (which centre on Gmail, Drive, Calendar, Fitness and
> Contacts), but Google's published list could not be retrieved to confirm it, and getting this wrong is
> costly. Check the current list before the OAuth verification submission, and treat **any** scope
> addition beyond these three as a plan change rather than a detail.

## Why the HTTP path is hand-written here

The other connectors share `@jobo-ai/connector-core`'s `JoboClient`. This one deliberately does not.

`JoboClient` is promise-based, and Apps Script has no event loop: a handler cannot return a promise across
the `google.script.run` boundary, and nothing guarantees the microtask queue drains before a synchronous
handler returns. Spin-waiting on a promise would work in testing and fail unpredictably in production.

So `Code.ts` makes blocking `UrlFetchApp` calls. The retry decision still comes from connector-core:
`request()` turns each failed response into the error the core client would throw (`toCoreError()`) and
asks the core's `isRetryable()` and `nextDelayMs()` whether to try again and how long to wait. That gives
the add-on the core's rules:

- The search (`POST /api/jobs/search`) and the signals read (`GET /api/signals`) bill every job or signal
  they return, so they are replayed only on a 429 or a 503 that carries `Retry-After`. After a dropped
  connection, a 500 or any other 5xx the charge may already have been made, so the import stops there
  instead.
- A 503 without `Retry-After` ("Search took too long") is never retried, on any request: the same search
  runs out of time again.
- The free requests (filter options, typeahead, occupations) are also retried after a dropped connection
  or a 5xx.
- `Retry-After` is honoured on any status, capped at 30 seconds.

It also shares the canonical filter values, the `Job` type, API key validation and `toJobSearchBody()`
(the sidebar request becomes a `POST /api/jobs/search` body through it), so those cannot drift.

## Imports and Apps Script's 6-minute limit

An Apps Script run is killed after 6 minutes. `importJobs()` and `importSignals()` write each page to the
sheet as it arrives (`sheetWriter()`: one `setValues` per page, header row with the first, then
`SpreadsheetApp.flush()`), so a page that was billed is in the sheet before the next one is requested. They
stop starting new pages, and skip any retry whose wait would end, after `IMPORT_TIME_BUDGET_MS` (4.5
minutes), and the message says how many were written and that the import stopped early.

## Filters

The sidebar covers connector-core's `JOB_SEARCH_FILTERS` (pinned `0.8.0`) except `posted_before`,
`discovered_after`, `discovered_before` (a one-shot import has no use for those windows) and
`description_format` (no description column). `SHEETS_SUPPORTED_FILTER_KEYS` plus the drift-guard test in
`Code.test.ts` keep that list honest.

| Picker | Source |
| --- | --- |
| Location | `suggestLocations` → `GET /api/connectors/locations/suggest`, per keystroke |
| Companies (include, exclude) | `suggestCompanies` → `GET /api/connectors/companies/suggest`, per keystroke; a pick sends every board id of the company |
| Occupations (include, exclude) | `getOccupations` → `GET /api/occupations` once per sidebar open, filtered locally (too large for `CacheService`) |
| Sources, industries, skills, job categories, job subcategories | `getFilterOptions` → `GET /api/connectors/filter-options/{facet}?limit=250`, script-cached 6 h, filtered locally |
| Languages, company categories, work model, employment type, experience level | Static lists from connector-core |

Every `importJobs` field is optional and an empty one is no filter, so an older sidebar's request still
works. Pages are equal-sized (`importPageSize`): `page` offsets by `page_size`, so shrinking the last page
used to re-fetch jobs already imported.

`muteHttpExceptions: true` is load-bearing: without it Apps Script throws on any non-2xx and discards the
response body, taking the problem details and the credit headers with it.

## Hiring Signals

**Import hiring signals…** opens the same sidebar as **Import jobs…**, on its signals view: one file, so the
two share the connection (sign-in, key, Disconnect). `Sidebar.html` is an Apps Script template for this one
value: `openSidebar()` prints the view into `data-view` on the body. Keep scriptlet tags out of the rest of
the file and of the inlined `connector-ui` assets; a test checks the built file has exactly one.

`importSignals()` reads `GET /api/signals` through the same `request()` as the job search, marked as billed
(see above). connector-core's `JoboClient` and `pollSignals` are async and cannot run here; its pure helpers
are used as they are: `toSignalsQuery()` builds the query (and refuses a value that contains a comma),
`describeSignal()`, `signalTypeLabel()` and `signalStrengthBand()` fill the Summary, Signal and Strength band
columns. A signal type the core does not know is written with its type in words.

- **Filters sent** (`SHEETS_SIGNAL_FILTER_KEYS`, checked against `HIRING_SIGNAL_FILTERS` by a test):
  `domains`, `types`, `countries`, `company_sizes`, `functions`, `min_strength`, `detected_after`,
  `exclude_agencies`, and `created_after` for "only new". The focus filters beyond `functions`, the
  thresholds and the funding filters stay in the dashboard and the API.
- **Domains** are cleaned like `HiringSignalDomains.Normalize` on the server (`cleanDomains()`): scheme,
  `www.`, path, user and port removed, lower-cased, de-duplicated. A typed value that is not a domain stops
  the import before anything is read; from the selected cells (`getSelectedDomains()`) it is left out and
  counted. Without domains, at least one of connector-core's `NARROWING_SIGNAL_FILTER_KEYS` must be set.
- **More than 100 domains** are read in chunks of `SIGNAL_MAX_LIST_VALUES` (100), one after another. The
  "how many" cap (1 to 1,000, default 50) counts signals *read*, which is what is billed, across all chunks;
  each request asks for no more than is left of it, at most 200 a page, so an import never reads past it.
- **Columns**: Detected, Company, Domain, Signal, Summary, Strength, Strength band, Country, Company size,
  Industry, Funding stage, Functions, Role, Seniority, Topics, Job countries, Evidence job IDs, Created at,
  Signal ID. Company size, Created at and Signal ID are set to plain text before they are written (Sheets
  reads `1-10` as a date, and would round `created_at`).
- **Only new** appends to the open tab when its header row starts with those headers (`readSignalsTab()`).
  It reads the stored order (`sort=created_asc`, oldest first) with `created_after` set to the tab's newest
  Created at, sent exactly as the tab holds it (`signalsReadBound()`). The API compares every decimal place,
  and a signal that becomes readable later never has an earlier `created_at`, so nothing in the tab is read
  again and nothing stored since is missed: there is no overlap and no waiting time. It needs an API that
  has the stored order. A signals tab with no rows yet is read like a new tab.
- An import into a new tab reads the newest detection day first, so one that stops early (cap, time,
  failure) can leave signals behind that are older than the newest one it wrote, and a later "only new"
  import never reads them. The message says so. An "only new" import that stops early is carried on exactly
  by the next one, because it wrote everything stored up to its last row; the message says that instead. The
  exception is more than 100 domains: the chunks are read one after another, so stopping early can leave a
  later chunk's signals behind, and the first message is shown.
- **Failures** (`describeSignalsFailure()`): 402 (nothing charged for that page), 403
  `account_key_required`, 404 `not_available_in_sandbox`, 429 (with the rate limits without a subscription),
  and a 400's first validation message.

Each signal costs ten times a job; see the [guide](https://jobo.world/docs/guides/hiring-signals).

## Storage

The API key is stored in **user** properties, not document properties. Spreadsheets get shared, and a
document property would hand the key to everyone it is shared with. A key that came from signing in is
stored the same way, with the account's email next to it so the sidebar can say who is connected.

## Sign in with Jobo

The sidebar's default way to connect. It is OAuth 2.0 authorization code with PKCE (S256) as a public
client, written by hand in `Code.ts` with no library. There is no client secret. The access
token that comes back is an ordinary Jobo API key: it goes through `saveApiKey()` like a pasted one, does
not expire, and is never refreshed. Pasting a key stays available under **Paste an API key instead**.

What Apps Script dictates:

- **The redirect address** is `https://script.google.com/macros/d/{SCRIPT_ID}/usercallback`, built at run
  time from `ScriptApp.getScriptId()`. The `state` value is an Apps Script state token that names the
  callback function (`authCallback`); Apps Script checks it before calling the function.
- **A sidebar cannot navigate the window**, so **Sign in with Jobo** is a link that opens a new tab. The
  sidebar then polls `getConnectionStatus()` and also offers **I have signed in**.
- **The PKCE verifier stays server-side**, in the user cache under an attempt id carried inside the state
  token, and is removed when the callback runs.

**Which client it signs in as.** The redirect address contains the script id, and Jobo compares it exactly.
Before each sign-in `resolveOAuthClientId()` asks Jobo whether this script's address is registered for the
client `google-sheets`:

- **It is:** the add-on signs in as `google-sheets`, which the consent screen shows as a verified app. That
  is the published add-on, once its address is set on the portal API as
  `OpenIddict:Connectors:google-sheets:RedirectUris:0`. The script id in that address is the `scriptId` in
  the `.clasp.json` that `clasp` writes for the deployed add-on project; it is not in this repository.
- **It is not:** this is a copy in its own script project (anyone installing from the mirror, a test
  deployment, or the published add-on before its address is registered). It registers itself once with
  Jobo's client registration endpoint, keeps the client id in script properties, and signs in as that
  client. The consent screen shows it as an unverified app named "Google Sheets (Career Site Jobs)" that
  returns to `script.google.com`.

So sign-in works for every copy without any server setting, and registering the published add-on's address
only changes how it is shown. The answer is cached for an hour, so the change takes that long to show.

`Disconnect` removes the key and the email from this add-on. After a sign-in it first hands the key back
to Jobo (`POST /connect/revoke`), which revokes it and removes the add-on from the connected apps under
**Developers → API Keys** in the Jobo dashboard. If Jobo cannot be reached, the sidebar says to disconnect
the add-on there. A pasted key is never sent: it is removed from the add-on and keeps working.

## Development

```bash
npm install && npm run build && npm test
```

`build.mjs` bundles everything (including connector-core) into a single `dist/Code.gs`, since Apps Script
has no module system, and emits a separate ESM copy under `dist-test/` so the pure helpers can be tested on
Node without an Apps Script runtime.

Push to a bound script:

```bash
npx --yes @google/clasp login && npm run push
```

## Not covered by the tests

The live API round trip, and anything requiring the Apps Script runtime itself — the real `UrlFetchApp`,
`PropertiesService`, sheet writes, and the sidebar. The tests cover retry classification, failure
messaging, query building, location formatting, sheet naming and the picker option mapping, plus the import
loop against a scripted fake of those services and a fake clock (the exact search body every filter
produces, equal page sizes, API warnings in the message, each page written before the next is requested,
the time budget, which failures are retried, and a failed later page keeping the pages already billed).
The signals import is covered the same way: domain cleaning, the query (chunks of 100 sharing one cap,
cursor paging), the rows and their plain-text columns, both "only new" modes, the result message and each
error message. They also check the built `dist/` (the manifest and its three scopes, the shared UI inlined
verbatim, and the sidebar's one scriptlet). Everything else needs a real deployment and a live key: a
`jbe_test_` key cannot read signals.

Not proven outside the fakes: `HtmlService.createTemplateFromFile` evaluating the whole sidebar with the
inlined combobox; `setNumberFormat("@")` keeping a seven-decimal `created_at` exactly as text;
`getActiveRange()` and `getActiveSheet()` from a sidebar call returning what the user has selected and
open; and the signals view's markup and script, which no test runs.

Sign in with Jobo is covered against scripted fakes only: the authorization URL, the challenge for a known
verifier, self-registration, and every outcome of the callback. It has never run in Apps Script. To check
in a real deployment, in this order: that `ScriptApp.newStateToken()` and `ScriptApp.getScriptId()` run
under the three scopes above (if Apps Script asks for another scope, that is a scope addition, see the
warning above); that script properties and the script cache are writable when the add-on runs for a user
who installed it from the Marketplace; that the add-on's `/usercallback` reaches `authCallback` for such a
user; and the full round trip, first as a self-registered copy and then with the address registered.

## Publishing

This is an **editor add-on**, not a Google Workspace add-on. It adds a menu under **Extensions**
(`onOpen`, and `onInstall` for the file open at install time) and shows an HTML sidebar. So
`appsscript.json` has no `addOns` section: that section, with its homepage trigger that must return Cards,
is only for Workspace add-ons, which are built with the Card service and cannot open an HTML sidebar. The
manifest holds just the runtime and the three scopes above. The add-on's name and logo are set in the
Marketplace SDK's store listing, not in the manifest.

Unlike the other connectors, Sheets has no CI publish step — `clasp login` and the Workspace Marketplace
console are both interactive, and `.clasp.json` is gitignored on purpose (it's a per-developer script
binding, not something to share). See [`../RELEASING.md`](../RELEASING.md) for the full sequence: `clasp
push`, the OAuth consent screen, the scope-classification check above, and the Marketplace listing review.
For an editor add-on, Google's publishing steps also need a numbered version of the script project
(create one after `clasp push`), whose number goes into the Marketplace SDK's app configuration as an
editor add-on for Sheets, and a standard Google Cloud project rather than the script's default one.

> ⚠️ Re-read the scope warning above before every OAuth verification submission — it is the single most
> expensive way this add-on can go wrong.
