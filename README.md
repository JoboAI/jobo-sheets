# Jobo for Sheets

Google Sheets add-on that imports filtered Jobo job data into the active spreadsheet.

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

So `Code.ts` makes blocking `UrlFetchApp` calls and reproduces the part of connector-core's behaviour that
actually matters — **which statuses are terminal** — explicitly, with tests. It still shares the canonical
filter values, the `Job` type, API key validation and `toJobSearchBody()` (the sidebar request becomes a
`POST /api/jobs/search` body through it), so those cannot drift.

## Filters

The sidebar covers connector-core's `JOB_SEARCH_FILTERS` (pinned `0.4.0`) except `posted_before`,
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

## Storage

The API key is stored in **user** properties, not document properties. Spreadsheets get shared, and a
document property would hand the key to everyone it is shared with.

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
loop against a scripted fake of those services (the exact search body every filter produces, equal page
sizes, API warnings in the message, and a failed later page still writing the pages already billed).
Everything else needs a real deployment and a `jbe_test_` key.

## Publishing

Unlike the other connectors, Sheets has no CI publish step — `clasp login` and the Workspace Marketplace
console are both interactive, and `.clasp.json` is gitignored on purpose (it's a per-developer script
binding, not something to share). See [`../RELEASING.md`](../RELEASING.md) for the full sequence: `clasp
push`, the OAuth consent screen, the scope-classification check above, and the Marketplace listing review.

> ⚠️ Re-read the scope warning above before every OAuth verification submission — it is the single most
> expensive way this add-on can go wrong.
