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
filter values, the `Job` type and API key validation, so those cannot drift.

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

The live API round trip, and anything requiring the Apps Script runtime itself — `UrlFetchApp`,
`PropertiesService`, sheet writes, and the sidebar. The tests cover retry classification, failure
messaging, query building, location formatting and sheet naming. Everything else needs a real deployment
and a `jbe_test_` key.

## Publishing

Unlike the other connectors, Sheets has no CI publish step — `clasp login` and the Workspace Marketplace
console are both interactive, and `.clasp.json` is gitignored on purpose (it's a per-developer script
binding, not something to share). See [`../RELEASING.md`](../RELEASING.md) for the full sequence: `clasp
push`, the OAuth consent screen, the scope-classification check above, and the Marketplace listing review.

> ⚠️ Re-read the scope warning above before every OAuth verification submission — it is the single most
> expensive way this add-on can go wrong.
