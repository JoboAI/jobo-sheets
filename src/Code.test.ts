import assert from "node:assert/strict";
import { createHash, randomUUID } from "node:crypto";
import { describe, it } from "node:test";

import {
  SHEETS_SUPPORTED_FILTER_KEYS,
  authCallback,
  buildQuery,
  codeChallenge,
  companyPickerItem,
  createCodeVerifier,
  describeFailure,
  disconnect,
  flattenOccupations,
  formatLocation,
  getConnectionStatus,
  importJobs,
  importPageSize,
  isRetryableStatus,
  neutralizeFormula,
  oauthRedirectUri,
  occupationLevelOptions,
  saveApiKey,
  startSignIn,
  toSearchFilters,
  uniqueSheetName,
} from "./Code.js";
import { JOB_SEARCH_FILTERS, type Job } from "@jobo-ai/connector-core";

function job(overrides: Partial<Job> = {}): Job {
  return {
    id: "job-1",
    title: "Engineer",
    normalized_title: null,
    company: null,
    description: null,
    summary: null,
    listing_url: null,
    apply_url: null,
    locations: [],
    compensation: null,
    employment_type: null,
    workplace_type: null,
    experience_level: null,
    source: "greenhouse",
    created_at: "2026-07-01T00:00:00Z",
    updated_at: "2026-07-01T00:00:00Z",
    date_posted: null,
    valid_through: null,
    qualifications: null,
    responsibilities: [],
    benefits: [],
    is_work_auth_required: null,
    is_h1b_sponsor: null,
    is_clearance_required: null,
    ...overrides,
  };
}

function jobs(count: number, offset = 0): Job[] {
  return Array.from({ length: count }, (_, n) => job({ id: `job-${offset + n}`, title: `Job ${offset + n}` }));
}

interface FakeResponse {
  status: number;
  body?: unknown;
  headers?: Record<string, string>;
}

/**
 * Just enough of the Apps Script runtime for importJobs: a saved key, a
 * scripted UrlFetchApp (an Error entry is thrown, like a network fault), and a
 * spreadsheet that records what was written.
 */
function installAppsScript(responses: Array<FakeResponse | Error>): {
  urls: string[];
  bodies: Array<Record<string, unknown> | null>;
  written: unknown[][][];
} {
  const urls: string[] = [];
  const bodies: Array<Record<string, unknown> | null> = [];
  const written: unknown[][][] = [];
  const g = globalThis as unknown as Record<string, unknown>;

  g.PropertiesService = { getUserProperties: () => ({ getProperty: () => "jbe_test_key" }) };
  g.Utilities = { sleep: () => undefined };
  g.UrlFetchApp = {
    fetch: (url: string, options: { method?: string; payload?: string } = {}) => {
      urls.push(url);
      bodies.push(options.payload ? (JSON.parse(options.payload) as Record<string, unknown>) : null);
      const next = responses.shift();
      if (!next) throw new Error(`unexpected request ${url}`);
      if (next instanceof Error) throw next;
      return {
        getResponseCode: () => next.status,
        getContentText: () => (next.body === undefined ? "" : JSON.stringify(next.body)),
        getAllHeaders: () => next.headers ?? {},
      };
    },
  };
  g.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheetByName: () => null,
      insertSheet: () => ({
        getRange: () => ({
          setValues: (rows: unknown[][]) => written.push(rows),
          setFontWeight: () => undefined,
        }),
        setFrozenRows: () => undefined,
        autoResizeColumns: () => undefined,
      }),
    }),
  };

  return { urls, bodies, written };
}

const billedPage = (batch: Job[], totalPages: number): FakeResponse => ({
  status: 200,
  body: { jobs: batch, total_pages: totalPages },
  headers: { "X-Credits-Deducted": String(batch.length * 3), "X-Credits-Balance": "40" },
});

describe("importJobs", () => {
  it("writes the pages already billed when a later page returns 402", () => {
    const { urls, bodies, written } = installAppsScript([
      billedPage(jobs(100), 3),
      { status: 402, body: { detail: "Balance too low." } },
    ]);

    const result = importJobs({ limit: 300 });

    assert.equal(written.length, 1, "page 1 should land in a sheet");
    assert.equal(written[0].length, 101, "header row + the 100 fetched jobs");
    assert.equal(result.ok, false);
    assert.match(result.message, /^Imported 100 of the 300 jobs requested before the import stopped\./);
    assert.match(result.message, /wallet balance is too low/i);
    assert.match(result.message, /Cost: \$0\.30 \(300 credits\)/);
    assert.equal(urls.length, 2, "a 402 is never retried");
    assert.equal(bodies[1]?.page, 2);
  });

  it("posts every sidebar filter to the search body and keeps the page size constant", () => {
    const { urls, bodies } = installAppsScript([billedPage(jobs(75), 2), billedPage(jobs(75, 75), 2)]);

    importJobs({
      q: "nurse",
      location: "Berlin, Germany",
      sources: ["greenhouse"],
      companies: ["id-1", "id-2", "acme.com"],
      excludeCompanies: ["FedEx, Inc."],
      skills: ["Python"],
      excludeSkills: ["PHP"],
      industries: ["Hospitals"],
      excludeIndustries: ["Staffing"],
      companyCategories: ["saas"],
      excludeCompanyCategories: ["agency"],
      jobCategories: ["Healthcare"],
      excludeJobCategories: ["Sales"],
      jobSubcategories: ["Nursing"],
      excludeJobSubcategories: ["Midwifery"],
      occupationIds: ["esco:abc"],
      excludeOccupationIds: ["jobo:backend-developer"],
      languages: ["de"],
      workModels: ["remote"],
      employmentTypes: ["full-time"],
      experienceLevels: ["senior"],
      minSalaryUsd: 50000,
      maxSalaryUsd: 0,
      postedAfter: "2026-11-01",
      searchDescriptions: false,
      sort: "discovered_desc",
      limit: 150,
    });

    assert.match(urls[0], /\/api\/jobs\/search$/);
    assert.deepEqual(bodies[0], {
      queries: ["nurse"],
      search_description: false,
      locations: ["Berlin, Germany"],
      sources: ["greenhouse"],
      work_models: ["remote"],
      employment_types: ["full-time"],
      experience_levels: ["senior"],
      languages: ["de"],
      skills: { include: ["Python"], exclude: ["PHP"] },
      companies: { include: ["id-1", "id-2", "acme.com"], exclude: ["FedEx, Inc."] },
      industries: { include: ["Hospitals"], exclude: ["Staffing"] },
      company_categories: { include: ["saas"], exclude: ["agency"] },
      job_categories: { include: ["Healthcare"], exclude: ["Sales"] },
      job_subcategories: { include: ["Nursing"], exclude: ["Midwifery"] },
      occupation_ids: { include: ["esco:abc"], exclude: ["jobo:backend-developer"] },
      salary_usd: { min: 50000 },
      posted_after: "2026-11-01",
      sort: "discovered_desc",
      page: 1,
      page_size: 75,
      include_facets: [],
      include_fields: [],
    });
    assert.equal(bodies[1]?.page, 2);
    assert.equal(bodies[1]?.page_size, 75, "page 2 must use page 1's size or it re-fetches page-1 jobs");
  });

  it("sends nothing but paging for an empty or older-shaped request", () => {
    const { bodies } = installAppsScript([billedPage(jobs(10), 1)]);

    importJobs({ q: "", workModels: [], companies: [], sort: "bogus", limit: 10 });

    assert.deepEqual(bodies[0], { page: 1, page_size: 10, include_facets: [], include_fields: [] });
  });

  it("passes the API's filter warnings on to the user", () => {
    installAppsScript([
      {
        status: 200,
        body: { jobs: [], total_pages: 0, warnings: [{ code: "companies_unmatched", message: "Bain matched no company." }] },
      },
    ]);

    const result = importJobs({ companies: ["Bain"], limit: 10 });

    assert.equal(result.ok, true);
    assert.equal(result.message, "No jobs matched those filters. Note: Bain matched no company.");
  });

  it("never writes more rows than requested when equal pages overshoot", () => {
    const { written } = installAppsScript([billedPage(jobs(51), 2), billedPage(jobs(51, 51), 2)]);

    importJobs({ limit: 101 });

    assert.equal(written[0].length, 102, "header + 101 jobs");
  });

  it("writes the pages already billed when a later page keeps failing after retries", () => {
    const { urls, written } = installAppsScript([
      billedPage(jobs(100), 2),
      { status: 503 },
      { status: 503 },
      { status: 503 },
    ]);

    const result = importJobs({ limit: 200 });

    assert.equal(urls.length, 4, "page 2 is tried three times");
    assert.equal(written[0].length, 101);
    assert.equal(result.ok, false);
    assert.match(result.message, /^Imported 100 of the 200 jobs requested/);
    assert.match(result.message, /HTTP 503/);
  });

  it("writes the pages already billed when a later fetch throws", () => {
    const { written } = installAppsScript([billedPage(jobs(100), 2), new Error("Address unavailable")]);

    const result = importJobs({ limit: 200 });

    assert.equal(written[0].length, 101);
    assert.equal(result.ok, false);
    assert.match(result.message, /Address unavailable/);
  });

  it("writes nothing when the first page fails", () => {
    const { written } = installAppsScript([{ status: 402, body: { detail: "Balance too low." } }]);

    const result = importJobs({ limit: 300 });

    assert.equal(written.length, 0);
    assert.equal(result.ok, false);
    assert.match(result.message, /^Your Jobo wallet balance is too low/);
  });

  it("reports a complete import as before", () => {
    const { written } = installAppsScript([billedPage(jobs(100), 2), billedPage(jobs(50, 100), 2)]);

    const result = importJobs({ limit: 150 });

    assert.equal(written[0].length, 151);
    assert.equal(result.ok, true);
    assert.match(result.message, /^Imported 150 jobs\. Cost: \$0\.45 \(450 credits\)\./);
  });
});

const LIVE_KEY = `jbe_live_${"a".repeat(21)}_${"b".repeat(43)}`;
const OTHER_KEY = `jbe_live_${"c".repeat(21)}_${"d".repeat(43)}`;
const SCRIPT_ID = "1AbCdEfGhIjKlMnOpQrStUvWxYz_0123456789-script";
const REDIRECT_URI = `https://script.google.com/macros/d/${SCRIPT_ID}/usercallback`;
const TOKEN_URL = "https://enterprise.jobo.world/connect/token";
const CLIENT_LOOKUP_URL = "https://enterprise.jobo.world/api/auth/oauth/client";
const REGISTER_URL = "https://enterprise.jobo.world/connect/register";

interface StateToken {
  token: string;
  method: string;
  timeout: number;
  args: Record<string, string>;
}

interface FetchCall {
  url: string;
  method?: string;
  contentType?: string;
  headers?: Record<string, string>;
  payload?: string;
  muteHttpExceptions?: boolean;
}

interface CallbackPage {
  file: string;
  title: string;
  ok?: boolean;
  heading?: string;
  message?: string;
  hint?: string;
}

/**
 * The Apps Script services Sign in with Jobo touches: user properties and the
 * user cache as maps, a state-token builder that records what it was given, a
 * SHA-256 digest that returns signed bytes and a padded web-safe base64 (both
 * as Apps Script does), a scripted UrlFetchApp, and the Callback template,
 * which records the values Code.ts hands it.
 */
function installSignIn(
  responses: Array<FakeResponse | Error> = [],
  stored: Record<string, string> = {},
  server: {
    /** What the client lookup answers. Default: Jobo has this script's address registered for google-sheets. */
    lookup?: FakeResponse | Error;
    /** What self-registration answers, in order. */
    registrations?: Array<FakeResponse | Error>;
    /** Script properties already saved by this script project. */
    scriptProperties?: Record<string, string>;
  } = {},
): {
  properties: Map<string, string>;
  scriptProperties: Map<string, string>;
  cache: Map<string, { value: string; seconds: number }>;
  stateTokens: StateToken[];
  fetches: FetchCall[];
  lookups: FetchCall[];
  registrations: FetchCall[];
  pages: CallbackPage[];
} {
  const properties = new Map(Object.entries(stored));
  const scriptProperties = new Map(Object.entries(server.scriptProperties ?? {}));
  const cache = new Map<string, { value: string; seconds: number }>();
  const scriptCache = new Map<string, string>();
  const stateTokens: StateToken[] = [];
  const fetches: FetchCall[] = [];
  const lookups: FetchCall[] = [];
  const registrations: FetchCall[] = [];
  const pendingRegistrations = [...(server.registrations ?? [])];
  const pages: CallbackPage[] = [];
  const g = globalThis as unknown as Record<string, unknown>;
  const store = (map: Map<string, string>) => ({
    getProperty: (key: string) => map.get(key) ?? null,
    setProperty: (key: string, value: string) => void map.set(key, value),
    deleteProperty: (key: string) => void map.delete(key),
  });

  g.PropertiesService = {
    getUserProperties: () => store(properties),
    getScriptProperties: () => store(scriptProperties),
  };
  g.CacheService = {
    getUserCache: () => ({
      get: (key: string) => cache.get(key)?.value ?? null,
      put: (key: string, value: string, seconds: number) => void cache.set(key, { value, seconds }),
      remove: (key: string) => void cache.delete(key),
    }),
    getScriptCache: () => ({
      get: (key: string) => scriptCache.get(key) ?? null,
      put: (key: string, value: string) => void scriptCache.set(key, value),
      remove: (key: string) => void scriptCache.delete(key),
    }),
  };
  g.Utilities = {
    sleep: () => undefined,
    getUuid: () => randomUUID(),
    DigestAlgorithm: { SHA_256: "SHA_256" },
    Charset: { US_ASCII: "US_ASCII" },
    computeDigest: (algorithm: string, value: string) => {
      assert.equal(algorithm, "SHA_256");
      return Array.from(createHash("sha256").update(value, "ascii").digest(), (byte) => (byte > 127 ? byte - 256 : byte));
    },
    base64EncodeWebSafe: (bytes: number[]) =>
      Buffer.from(bytes.map((byte) => byte & 0xff))
        .toString("base64")
        .replace(/\+/g, "-")
        .replace(/\//g, "_"),
  };
  g.ScriptApp = {
    getScriptId: () => SCRIPT_ID,
    newStateToken: () => {
      // Apps Script's defaults: a function named callback, valid for 60 seconds.
      const pending: StateToken = { token: `state-${stateTokens.length + 1}`, method: "callback", timeout: 60, args: {} };
      const builder = {
        withMethod: (method: string) => ((pending.method = method), builder),
        withArgument: (name: string, value: string) => ((pending.args[name] = value), builder),
        withTimeout: (seconds: number) => ((pending.timeout = seconds), builder),
        createToken: () => (stateTokens.push(pending), pending.token),
      };
      return builder;
    },
  };
  g.HtmlService = {
    createTemplateFromFile: (file: string) => {
      const page: CallbackPage = { file, title: "" };
      pages.push(page);
      const output = { setTitle: (title: string) => ((page.title = title), output) };
      return Object.assign(page, { evaluate: () => output });
    },
  };
  const answer = (next: FakeResponse | Error | undefined, url: string) => {
    if (!next) throw new Error(`unexpected request ${url}`);
    if (next instanceof Error) throw next;
    return {
      getResponseCode: () => next.status,
      getContentText: () => (next.body === undefined ? "" : JSON.stringify(next.body)),
      getAllHeaders: () => next.headers ?? {},
    };
  };
  g.UrlFetchApp = {
    fetch: (url: string, options: Omit<FetchCall, "url"> = {}) => {
      // Which client to sign in as is settled before the sign-in itself: kept
      // apart so `fetches` stays the token requests.
      if (url.startsWith(CLIENT_LOOKUP_URL)) {
        lookups.push({ url, ...options });
        return answer(server.lookup ?? { status: 200, body: { client_id: "google-sheets", redirect_host: "script.google.com" } }, url);
      }
      if (url === REGISTER_URL) {
        registrations.push({ url, ...options });
        return answer(pendingRegistrations.shift(), url);
      }
      fetches.push({ url, ...options });
      return answer(responses.shift(), url);
    },
  };

  return { properties, scriptProperties, cache, stateTokens, fetches, lookups, registrations, pages };
}

/** What Apps Script passes the callback: the redirect's query plus the state token's arguments. */
function callbackRequest(state: StateToken, query: Record<string, string>): { parameter: Record<string, string> } {
  return { parameter: { ...query, state: state.token, ...state.args } };
}

const tokenResponse = (body: Record<string, unknown>): FakeResponse => ({ status: 200, body });

describe("sign in with Jobo", () => {
  it("builds the authorization URL from the script id, with S256 and no verifier in it", () => {
    const { cache, stateTokens } = installSignIn();

    const { url } = startSignIn();

    const parsed = new URL(url);
    assert.equal(parsed.origin + parsed.pathname, "https://enterprise.jobo.world/connect/authorize");
    const [[cacheKey, kept]] = [...cache];
    assert.deepEqual(Object.fromEntries(parsed.searchParams), {
      response_type: "code",
      client_id: "google-sheets",
      redirect_uri: REDIRECT_URI,
      scope: "api email",
      state: "state-1",
      code_challenge: createHash("sha256").update(kept.value).digest("base64url"),
      code_challenge_method: "S256",
    });
    assert.match(url, /[?&]scope=api%20email&/, "the two scopes are space separated");
    assert.equal(oauthRedirectUri(), REDIRECT_URI);

    assert.ok(!url.includes(kept.value), "the verifier stays on the server");
    assert.equal(cacheKey, `jobo_oauth_verifier_${stateTokens[0].args.attempt}`);
    assert.equal(kept.seconds, 3600);
    assert.equal(stateTokens[0].timeout, 3600, "the link and its verifier expire together");
  });

  it("names a callback that Apps Script can find as a global", () => {
    const { stateTokens } = installSignIn();

    startSignIn();

    assert.equal(stateTokens[0].method, "authCallback");
    assert.equal((globalThis as unknown as Record<string, unknown>)[stateTokens[0].method], authCallback);
  });

  it("derives the challenge of RFC 7636's example verifier, unpadded", () => {
    installSignIn();

    assert.equal(
      codeChallenge("dBjftJeZ4CVP-mB92K27uhbUJU1p1r_wW1gFWFOEjXk"),
      "E9Melhoa2OwvFrEMTJguCHaoeK1t8URWbuGJSstw-cM",
    );
  });

  it("makes a new verifier of 43 to 128 unreserved characters for every attempt", () => {
    const { cache, stateTokens } = installSignIn();

    const first = createCodeVerifier();
    assert.match(first, /^[A-Za-z0-9\-._~]{43,128}$/);
    assert.notEqual(createCodeVerifier(), first);

    startSignIn();
    startSignIn();
    const kept = [...cache.values()].map((entry) => entry.value);
    assert.equal(kept.length, 2, "an earlier link keeps working after a newer one is issued");
    assert.notEqual(kept[0], kept[1]);
    assert.notEqual(stateTokens[0].args.attempt, stateTokens[1].args.attempt);
  });

  it("exchanges the code without a client secret and stores the key and the email", () => {
    const { properties, cache, stateTokens, fetches, pages } = installSignIn([
      tokenResponse({ access_token: LIVE_KEY, token_type: "Bearer", scope: "api", account_email: "user@example.com" }),
    ]);
    startSignIn();
    const verifier = [...cache.values()][0].value;

    authCallback(callbackRequest(stateTokens[0], { code: "the-code" }));

    assert.equal(fetches.length, 1);
    assert.equal(fetches[0].url, TOKEN_URL);
    assert.equal(fetches[0].method, "post");
    assert.equal(fetches[0].contentType, "application/x-www-form-urlencoded");
    assert.equal(fetches[0].muteHttpExceptions, true);
    assert.deepEqual(Object.fromEntries(new URLSearchParams(fetches[0].payload)), {
      grant_type: "authorization_code",
      client_id: "google-sheets",
      redirect_uri: REDIRECT_URI,
      code: "the-code",
      code_verifier: verifier,
    });
    assert.ok(!JSON.stringify(fetches[0]).includes("client_secret"), "a public client never sends a secret");
    assert.equal(fetches[0].headers?.Authorization, undefined);

    assert.equal(properties.get("JOBO_API_KEY"), LIVE_KEY);
    assert.equal(properties.get("JOBO_ACCOUNT_EMAIL"), "user@example.com");
    assert.equal(properties.get("JOBO_OAUTH_CLIENT"), "google-sheets", "who to hand the key back as");
    assert.equal(cache.size, 0, "the verifier is removed once used");
    assert.deepEqual(getConnectionStatus(), { connected: true, email: "user@example.com" });

    assert.equal(pages.length, 1);
    assert.equal(pages[0].file, "Callback");
    assert.equal(pages[0].ok, true);
    assert.equal(pages[0].heading, "Connected to Jobo");
    assert.equal(pages[0].message, "Your Jobo account user@example.com is connected.");
    assert.match(pages[0].hint ?? "", /close this tab/);
    assert.ok(!JSON.stringify(pages[0]).includes(LIVE_KEY), "the page never shows the key");
  });

  it("stores the signed-in key where a pasted key goes, so an import sends it unchanged", () => {
    const signedIn = installSignIn([
      tokenResponse({ access_token: LIVE_KEY, token_type: "Bearer", scope: "api" }),
      { status: 200, body: { jobs: [], total_pages: 0 } },
    ]);
    startSignIn();
    authCallback(callbackRequest(signedIn.stateTokens[0], { code: "the-code" }));
    const result = importJobs({ limit: 10 });

    assert.equal(result.message, "No jobs matched those filters.");
    assert.equal(signedIn.fetches[1].headers?.["X-Api-Key"], LIVE_KEY);
    assert.deepEqual(getConnectionStatus(), { connected: true, email: null }, "no email without the email scope");
    assert.equal(signedIn.pages[0].message, "Your Jobo account is connected.");

    // The key sits where a pasted key does. A sign-in also remembers the app it was issued to.
    const signedInProperties = new Map(signedIn.properties);
    assert.equal(signedInProperties.get("JOBO_OAUTH_CLIENT"), "google-sheets");
    signedInProperties.delete("JOBO_OAUTH_CLIENT");
    const pasted = installSignIn();
    assert.equal(saveApiKey(` ${LIVE_KEY} `).ok, true);
    assert.deepEqual([...pasted.properties], [...signedInProperties]);
  });

  it("uses each attempt's own verifier and leaves other attempts waiting", () => {
    const { cache, stateTokens, fetches } = installSignIn([tokenResponse({ access_token: LIVE_KEY })]);
    startSignIn();
    startSignIn();
    const second = stateTokens[1].args.attempt;
    const secondVerifier = cache.get(`jobo_oauth_verifier_${second}`)?.value;

    authCallback(callbackRequest(stateTokens[1], { code: "the-code" }));

    assert.equal(new URLSearchParams(fetches[0].payload).get("code_verifier"), secondVerifier);
    assert.deepEqual([...cache.keys()], [`jobo_oauth_verifier_${stateTokens[0].args.attempt}`]);
  });

  it("says sign-in was cancelled when the user declines, and saves nothing", () => {
    const { properties, cache, stateTokens, fetches, pages } = installSignIn();
    startSignIn();

    authCallback(
      callbackRequest(stateTokens[0], {
        error: "access_denied",
        error_description: "The authorization was denied by the end user.",
      }),
    );

    assert.equal(fetches.length, 0);
    assert.equal(properties.size, 0);
    assert.equal(cache.size, 0, "a declined attempt is over");
    assert.equal(pages[0].ok, false);
    assert.equal(pages[0].heading, "Sign-in was cancelled");
    assert.equal(pages[0].message, "The authorization was denied by the end user.");
    assert.deepEqual(getConnectionStatus(), { connected: false, email: null });
  });

  it("reports any other authorization error as a failed sign-in", () => {
    const { stateTokens, fetches, pages } = installSignIn();
    startSignIn();

    authCallback(callbackRequest(stateTokens[0], { error: "invalid_scope" }));

    assert.equal(fetches.length, 0);
    assert.equal(pages[0].ok, false);
    assert.equal(pages[0].heading, "Sign-in failed");
    assert.equal(pages[0].message, "Jobo returned the error invalid_scope.");
  });

  it("shows the server's reason when the token endpoint answers 400", () => {
    const { properties, cache, stateTokens, pages } = installSignIn([
      { status: 400, body: { error: "invalid_grant", error_description: "The authorization code is no longer valid." } },
    ]);
    startSignIn();

    authCallback(callbackRequest(stateTokens[0], { code: "used-code" }));

    assert.equal(properties.size, 0);
    assert.equal(cache.size, 0);
    assert.equal(pages[0].ok, false);
    assert.equal(pages[0].heading, "Sign-in failed");
    assert.equal(pages[0].message, "The authorization code is no longer valid.");
  });

  it("falls back to the status when a failed exchange has no description", () => {
    const { stateTokens, pages } = installSignIn([{ status: 502 }]);
    startSignIn();

    authCallback(callbackRequest(stateTokens[0], { code: "the-code" }));

    assert.equal(pages[0].ok, false);
    assert.equal(pages[0].message, "Jobo returned HTTP 502.");
  });

  it("refuses a token that is not a Jobo key and keeps it off the page", () => {
    const notAKey = "eyJhbGciOiJSUzI1NiJ9.not-a-jobo-key";
    const { properties, stateTokens, pages } = installSignIn(
      [tokenResponse({ access_token: notAKey, account_email: "user@example.com" }), tokenResponse({ token_type: "Bearer" })],
      { JOBO_API_KEY: OTHER_KEY },
    );

    startSignIn();
    authCallback(callbackRequest(stateTokens[0], { code: "the-code" }));
    startSignIn();
    authCallback(callbackRequest(stateTokens[1], { code: "another-code" }));

    assert.deepEqual([...properties], [["JOBO_API_KEY", OTHER_KEY]], "nothing is stored, and no email is attached");
    for (const page of pages) {
      assert.equal(page.ok, false);
      assert.equal(page.message, "Jobo did not return a valid API key.");
      assert.ok(!JSON.stringify(page).includes(notAKey));
    }
  });

  it("fails without calling the token endpoint when the attempt has no verifier", () => {
    const { properties, stateTokens, fetches, pages } = installSignIn([tokenResponse({ access_token: LIVE_KEY })]);
    startSignIn();

    // An attempt this user never started, and a callback with no attempt at all.
    authCallback({ parameter: { code: "the-code", state: "state-1", attempt: "someone-elses-attempt" } });
    authCallback({ parameter: { code: "the-code" } });

    assert.equal(fetches.length, 0);
    assert.equal(properties.size, 0);
    for (const page of pages) {
      assert.equal(page.ok, false);
      assert.equal(page.message, "This sign-in link has expired or was already used.");
    }

    // The real attempt still works once, and only once.
    authCallback(callbackRequest(stateTokens[0], { code: "the-code" }));
    authCallback(callbackRequest(stateTokens[0], { code: "the-code" }));

    assert.equal(fetches.length, 1, "a replayed callback never reaches the token endpoint");
    assert.equal(pages[2].ok, true);
    assert.equal(pages[3].ok, false);
    assert.equal(properties.get("JOBO_API_KEY"), LIVE_KEY);
  });

  it("fails when the redirect carries neither a code nor an error", () => {
    const { cache, stateTokens, fetches, pages } = installSignIn();
    startSignIn();

    authCallback(callbackRequest(stateTokens[0], {}));

    assert.equal(fetches.length, 0);
    assert.equal(cache.size, 0);
    assert.equal(pages[0].message, "Jobo did not return an authorization code.");
  });

  it("fails cleanly when the token endpoint cannot be reached", () => {
    const { properties, stateTokens, pages } = installSignIn([new Error("Address unavailable")]);
    startSignIn();

    authCallback(callbackRequest(stateTokens[0], { code: "the-code" }));

    assert.equal(properties.size, 0);
    assert.equal(pages[0].ok, false);
    assert.equal(pages[0].message, "Jobo could not be reached to finish signing in.");
  });

  describe("a copy of the add-on in its own script project", () => {
    const unknownToJobo: FakeResponse = { status: 404 };
    const registered = (clientId: string): FakeResponse => ({ status: 201, body: { client_id: clientId } });

    it("asks Jobo whether this script's address belongs to the published add-on", () => {
      const { lookups, registrations, stateTokens } = installSignIn();

      const { url } = startSignIn();

      assert.equal(lookups.length, 1);
      const lookup = new URL(lookups[0].url);
      assert.equal(lookup.searchParams.get("client_id"), "google-sheets");
      assert.equal(lookup.searchParams.get("redirect_uri"), REDIRECT_URI);
      assert.equal(registrations.length, 0, "the published add-on never registers itself");
      assert.equal(new URL(url).searchParams.get("client_id"), "google-sheets");
      assert.equal(stateTokens[0].args.client, "google-sheets");
    });

    it("registers itself once, signs in as that client, and exchanges the code as it", () => {
      const { lookups, registrations, scriptProperties, stateTokens, fetches, properties } = installSignIn(
        [tokenResponse({ access_token: LIVE_KEY, token_type: "Bearer", scope: "api" })],
        {},
        { lookup: unknownToJobo, registrations: [registered("self-123")] },
      );

      const first = startSignIn();
      const second = startSignIn();

      assert.equal(registrations.length, 1, "one registration per script project");
      assert.equal(lookups.length, 1, "and the answer is remembered between sign-in links");
      assert.equal(registrations[0].method, "post");
      assert.equal(registrations[0].contentType, "application/json");
      assert.deepEqual(JSON.parse(registrations[0].payload ?? ""), {
        client_name: "Google Sheets (Career Site Jobs)",
        redirect_uris: [REDIRECT_URI],
      });
      assert.deepEqual(JSON.parse(scriptProperties.get("JOBO_OAUTH_SELF_CLIENT") ?? ""), {
        clientId: "self-123",
        redirectUri: REDIRECT_URI,
      });
      for (const { url } of [first, second]) assert.equal(new URL(url).searchParams.get("client_id"), "self-123");
      assert.equal(stateTokens[0].args.client, "self-123");

      authCallback(callbackRequest(stateTokens[0], { code: "the-code" }));

      assert.equal(new URLSearchParams(fetches[0].payload).get("client_id"), "self-123");
      assert.ok(!JSON.stringify(fetches[0]).includes("client_secret"));
      assert.equal(properties.get("JOBO_API_KEY"), LIVE_KEY);
    });

    it("reuses the client it registered before, so signing in again returns the same key", () => {
      const { registrations } = installSignIn([], {}, {
        lookup: unknownToJobo,
        scriptProperties: { JOBO_OAUTH_SELF_CLIENT: JSON.stringify({ clientId: "self-old", redirectUri: REDIRECT_URI }) },
      });

      assert.equal(new URL(startSignIn().url).searchParams.get("client_id"), "self-old");
      assert.equal(registrations.length, 0);
    });

    it("registers again when the saved client belongs to another address", () => {
      const { registrations, scriptProperties } = installSignIn([], {}, {
        lookup: unknownToJobo,
        registrations: [registered("self-new")],
        scriptProperties: {
          JOBO_OAUTH_SELF_CLIENT: JSON.stringify({ clientId: "self-old", redirectUri: "https://script.google.com/macros/d/OTHER/usercallback" }),
        },
      });

      assert.equal(new URL(startSignIn().url).searchParams.get("client_id"), "self-new");
      assert.equal(registrations.length, 1);
      assert.equal(JSON.parse(scriptProperties.get("JOBO_OAUTH_SELF_CLIENT") ?? "").clientId, "self-new");
    });

    it("treats a client that exists without this address as not this add-on", () => {
      const { registrations } = installSignIn([], {}, {
        // google-sheets is registered, for a different script id: no redirect host is echoed back.
        lookup: { status: 200, body: { client_id: "google-sheets", redirect_host: null } },
        registrations: [registered("self-456")],
      });

      assert.equal(new URL(startSignIn().url).searchParams.get("client_id"), "self-456");
      assert.equal(registrations.length, 1);
    });

    it("does not start a sign-in it could not set up, and saves nothing", () => {
      for (const server of [
        { lookup: new Error("DNS") },
        { lookup: { status: 503 } as FakeResponse },
        { lookup: unknownToJobo, registrations: [new Error("timeout")] },
        { lookup: unknownToJobo, registrations: [{ status: 400, body: { error: "invalid_redirect_uri" } } as FakeResponse] },
        { lookup: unknownToJobo, registrations: [{ status: 201, body: {} } as FakeResponse] },
      ]) {
        const { scriptProperties, stateTokens, cache } = installSignIn([], {}, server);

        assert.throws(() => startSignIn(), /Jobo could not be reached to start signing in/);
        assert.equal(scriptProperties.size, 0);
        assert.equal(stateTokens.length, 0);
        assert.equal(cache.size, 0);
      }
    });
  });

  it("disconnect hands a signed-in key back to Jobo, then forgets it", () => {
    const { properties, fetches } = installSignIn([{ status: 200 }], {
      JOBO_API_KEY: LIVE_KEY,
      JOBO_ACCOUNT_EMAIL: "user@example.com",
      JOBO_OAUTH_CLIENT: "self-123",
    });
    assert.deepEqual(getConnectionStatus(), { connected: true, email: "user@example.com" });

    const result = disconnect();

    assert.equal(fetches.length, 1);
    assert.equal(fetches[0].url, "https://enterprise.jobo.world/connect/revoke");
    assert.equal(fetches[0].method, "post");
    assert.equal(fetches[0].contentType, "application/x-www-form-urlencoded");
    assert.deepEqual(Object.fromEntries(new URLSearchParams(fetches[0].payload)), {
      token: LIVE_KEY,
      client_id: "self-123",
    });
    assert.deepEqual(result, { revoked: true, message: "Disconnected. The key this add-on held has been revoked." });
    assert.equal(properties.size, 0);
    assert.deepEqual(getConnectionStatus(), { connected: false, email: null });
  });

  it("disconnect still forgets the key when Jobo cannot be reached, and says where to finish", () => {
    for (const answer of [new Error("Address unavailable"), { status: 503 } as FakeResponse]) {
      const { properties } = installSignIn([answer], { JOBO_API_KEY: LIVE_KEY, JOBO_OAUTH_CLIENT: "google-sheets" });

      const result = disconnect();

      assert.equal(result.revoked, false);
      assert.match(result.message, /could not be reached.*Developers → API Keys/);
      assert.ok(!result.message.includes(LIVE_KEY));
      assert.equal(properties.size, 0);
    }
  });

  it("disconnect never sends a pasted key anywhere", () => {
    const { properties, fetches } = installSignIn([], { JOBO_API_KEY: LIVE_KEY });

    assert.deepEqual(disconnect(), { revoked: true, message: "" });

    assert.equal(fetches.length, 0, "a pasted key is the user's own and stays live at Jobo");
    assert.equal(properties.size, 0);
  });

  it("a pasted key drops the previous account's email, and a rejected paste changes nothing", () => {
    const { properties } = installSignIn([], { JOBO_API_KEY: LIVE_KEY, JOBO_ACCOUNT_EMAIL: "user@example.com" });

    assert.equal(saveApiKey("not-a-key").ok, false);
    assert.deepEqual(getConnectionStatus(), { connected: true, email: "user@example.com" });

    assert.equal(saveApiKey(OTHER_KEY).ok, true);
    assert.deepEqual([...properties], [["JOBO_API_KEY", OTHER_KEY]]);
    assert.deepEqual(getConnectionStatus(), { connected: true, email: null });
  });

  it("a key pasted over a signed-in one is no longer handed back on disconnect", () => {
    const { properties, fetches } = installSignIn([], {
      JOBO_API_KEY: LIVE_KEY,
      JOBO_ACCOUNT_EMAIL: "user@example.com",
      JOBO_OAUTH_CLIENT: "google-sheets",
    });

    assert.equal(saveApiKey(OTHER_KEY).ok, true);
    assert.deepEqual([...properties], [["JOBO_API_KEY", OTHER_KEY]]);

    disconnect();
    assert.equal(fetches.length, 0);
  });
});

describe("retry policy", () => {
  it("never retries insufficient credits — the precheck would fail identically", () => {
    assert.equal(isRetryableStatus(402, null), false);
  });

  it("never retries a void feed cursor", () => {
    assert.equal(isRetryableStatus(409, "feed_cursor_restart_required"), false);
  });

  it("never retries an auth failure", () => {
    assert.equal(isRetryableStatus(401, null), false);
    assert.equal(isRetryableStatus(403, null), false);
  });

  it("retries throttling and server faults", () => {
    assert.equal(isRetryableStatus(429, null), true);
    assert.equal(isRetryableStatus(500, null), true);
    assert.equal(isRetryableStatus(503, null), true);
    assert.equal(isRetryableStatus(0, null), true);
  });

  it("does not retry a bad request", () => {
    assert.equal(isRetryableStatus(400, null), false);
    assert.equal(isRetryableStatus(404, null), false);
  });
});

describe("failure messages", () => {
  it("tells the user what to do about an empty wallet", () => {
    const message = describeFailure(402, { detail: "Balance too low." });
    assert.match(message, /wallet balance/i);
    assert.match(message, /enterprise\.jobo\.world/);
  });

  it("points at the key page on an auth failure", () => {
    assert.match(describeFailure(401, {}), /api-keys/);
  });

  it("falls back to the upstream detail", () => {
    assert.equal(describeFailure(400, { detail: "page_size must be between 1 and 100." }), "page_size must be between 1 and 100.");
  });

  it("still says something useful with no body", () => {
    assert.match(describeFailure(500, null), /HTTP 500/);
  });
});

describe("query building", () => {
  it("drops empty and undefined values", () => {
    assert.equal(buildQuery({ q: "rust", location: "", page: undefined }), "?q=rust");
  });

  it("encodes values", () => {
    assert.equal(buildQuery({ location: "Berlin, Germany" }), "?location=Berlin%2C%20Germany");
  });

  it("returns an empty string when nothing is set", () => {
    assert.equal(buildQuery({ q: "", page: undefined }), "");
  });

  it("carries the newer filters and omits them when empty", () => {
    const query = buildQuery({
      skills: "Python,Kubernetes",
      industries: "fintech",
      max_salary_usd: 150000,
      posted_after: "2026-07-01",
      search_description: "false",
    });
    assert.match(query, /skills=Python%2CKubernetes/);
    assert.match(query, /industries=fintech/);
    assert.match(query, /max_salary_usd=150000/);
    assert.match(query, /posted_after=2026-07-01/);
    assert.match(query, /search_description=false/);

    assert.equal(buildQuery({ skills: "", industries: "", posted_after: "" }), "");
  });
});

describe("filter drift guard", () => {
  it("covers every core search filter except the documented one-shot gaps", () => {
    // This connector is a one-shot import, not an incremental sync — the
    // backward-looking windows stay out deliberately, and the sheet has no
    // description column to format (see Code.ts).
    const knownGaps = new Set(["posted_before", "discovered_after", "discovered_before", "description_format"]);

    const coreKeys = JOB_SEARCH_FILTERS.map((filter) => filter.key);
    for (const key of coreKeys) {
      const covered = SHEETS_SUPPORTED_FILTER_KEYS.includes(key) || knownGaps.has(key);
      assert.ok(covered, `core filter "${key}" is neither supported nor a documented gap`);
    }

    for (const key of SHEETS_SUPPORTED_FILTER_KEYS) {
      assert.ok(coreKeys.includes(key), `"${key}" is not a core search filter`);
    }

    for (const gap of knownGaps) {
      assert.ok(
        !SHEETS_SUPPORTED_FILTER_KEYS.includes(gap),
        `"${gap}" is listed as both supported and a gap`,
      );
    }
  });
});

describe("picker options", () => {
  it("splits an import into the fewest equal pages of at most 100", () => {
    assert.equal(importPageSize(10), 10);
    assert.equal(importPageSize(100), 100);
    assert.equal(importPageSize(101), 51);
    assert.equal(importPageSize(150), 75);
    assert.equal(importPageSize(1000), 100);
  });

  it("turns a company suggestion into one chip carrying every board id", () => {
    const item = companyPickerItem({
      id: "b",
      ids: ["a", "b"],
      name: "Stripe",
      domain: "stripe.com",
      logo_url: null,
      country: "US",
      job_count: 43,
    });
    assert.deepEqual(item, { label: "Stripe (stripe.com, US)", value: "a,b", ids: ["a", "b"], count: 43 });

    const bare = companyPickerItem({ id: "c", ids: [], name: "Acme", domain: null, logo_url: null, country: null, job_count: null });
    assert.deepEqual(bare, { label: "Acme", value: "c", ids: ["c"] });
  });

  it("labels subcategories with their category and keeps the display name as the value", () => {
    const options = occupationLevelOptions(
      [
        { key: "nursing", count: 12, name: "Nursing", parent: "healthcare" },
        { key: "software development", count: 40 },
      ],
      true,
    );
    assert.deepEqual(options, [
      { value: "Nursing", label: "Nursing · Healthcare", count: 12 },
      { value: "Software Development", label: "Software Development", count: 40 },
    ]);
  });

  it("flattens the taxonomy into occupation ids labelled with their place", () => {
    const options = flattenOccupations({
      version: "1",
      categories: [
        {
          name: "Healthcare",
          subcategories: [{ name: "Nursing", occupations: [{ id: "esco:rn", name: "Registered Nurse" }] }],
        },
      ],
    });
    assert.deepEqual(options, [{ value: "esco:rn", label: "Registered Nurse (Healthcare › Nursing)" }]);
    assert.deepEqual(flattenOccupations(null), []);
  });

  it("maps the sidebar request onto the shared filter keys", () => {
    const filters = toSearchFilters({ languages: ["en"], occupationIds: ["esco:abc"], searchDescriptions: true, sort: "relevance" });
    assert.deepEqual(filters.language, ["en"]);
    assert.deepEqual(filters.occupation_ids, ["esco:abc"]);
    assert.equal(filters.search_description, undefined, "true is the default and is not sent");
    assert.equal(filters.sort, "relevance");
  });
});

describe("location formatting", () => {
  it("prefers the pre-rendered label", () => {
    const j = job({ locations: [{ location: "Remote, EU", city: null, region: null, country: null, latitude: null, longitude: null }] });
    assert.equal(formatLocation(j), "Remote, EU");
  });

  it("composes from parts when there is no label", () => {
    const j = job({ locations: [{ location: null, city: "Berlin", region: null, country: "Germany", latitude: null, longitude: null }] });
    assert.equal(formatLocation(j), "Berlin, Germany");
  });

  it("falls back to Remote when there is no location at all", () => {
    assert.equal(formatLocation(job({ workplace_type: "Remote" })), "Remote");
  });

  it("returns empty rather than a stray comma", () => {
    assert.equal(formatLocation(job()), "");
  });
});

describe("sheet naming", () => {
  it("uses the base name when it is free", () => {
    assert.equal(uniqueSheetName({ getSheetByName: () => null }, "Jobo Jobs"), "Jobo Jobs");
  });

  it("suffixes rather than overwriting an existing sheet", () => {
    const taken = new Set(["Jobo Jobs", "Jobo Jobs 2"]);
    const sheet = { getSheetByName: (name: string) => (taken.has(name) ? {} : null) };
    assert.equal(uniqueSheetName(sheet, "Jobo Jobs"), "Jobo Jobs 3");
  });
});

describe("neutralizeFormula", () => {
  it("prefixes text that Sheets would run as a formula", () => {
    assert.equal(neutralizeFormula('=HYPERLINK("https://evil.test","Apply")'), `'=HYPERLINK("https://evil.test","Apply")`);
    assert.equal(neutralizeFormula("=IMPORTDATA(A1)"), "'=IMPORTDATA(A1)");
    assert.equal(neutralizeFormula("+1 555 0100"), "'+1 555 0100");
    assert.equal(neutralizeFormula("-x"), "'-x");
    assert.equal(neutralizeFormula("@SUM(A1)"), "'@SUM(A1)");
    assert.equal(neutralizeFormula("\t=1"), "'\t=1");
  });

  it("leaves ordinary text and numbers alone", () => {
    assert.equal(neutralizeFormula("Engineer = great"), "Engineer = great");
    assert.equal(neutralizeFormula(""), "");
    assert.equal(neutralizeFormula(-5), -5);
    assert.equal(neutralizeFormula(120000), 120000);
  });
});
