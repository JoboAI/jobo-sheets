import assert from "node:assert/strict";
import { createHash, randomUUID } from "node:crypto";
import { readFileSync } from "node:fs";
import { afterEach, describe, it } from "node:test";

import {
  SHEETS_SIGNAL_FILTER_KEYS,
  SHEETS_SUPPORTED_FILTER_KEYS,
  SIGNAL_COLUMNS,
  SIGNAL_TYPE_GROUPS,
  authCallback,
  buildQuery,
  cleanDomains,
  codeChallenge,
  companyPickerItem,
  createCodeVerifier,
  describeFailure,
  describeSignalsFailure,
  disconnect,
  flattenOccupations,
  formatLocation,
  getConnectionStatus,
  getSelectedDomains,
  getSettings,
  getSignalsTabStatus,
  importJobs,
  importPageSize,
  importSignals,
  isRetryableResponse,
  neutralizeFormula,
  newestCreatedAt,
  oauthRedirectUri,
  occupationLevelOptions,
  onOpen,
  saveApiKey,
  showSidebar,
  showSignalsSidebar,
  signalTypeGroups,
  signalsReadBound,
  startSignIn,
  toCoreError,
  toSearchFilters,
  uniqueSheetName,
} from "./Code.js";
import {
  FeedCursorRestartRequiredError,
  HIRING_SIGNAL_FILTERS,
  InsufficientCreditsError,
  JOB_SEARCH_FILTERS,
  RateLimitError,
  SIGNAL_TYPES,
  SIGNAL_TYPE_LABELS,
  type HiringSignal,
  type Job,
} from "@jobo-ai/connector-core";

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
  /** How long the fake clock advances while this answer is on its way. */
  takesMs?: number;
}

const realDateNow = Date.now;
afterEach(() => {
  Date.now = realDateNow;
});

/** A tab in the fake spreadsheet: its name and its cells, row 1 first. */
interface FakeTab {
  name: string;
  grid: unknown[][];
}

/**
 * Just enough of the Apps Script runtime for the imports: a saved key, a
 * scripted UrlFetchApp (an Error entry is thrown, like a network fault), a
 * spreadsheet whose tabs are grids and which records each setValues and
 * setNumberFormat call, and a clock that only moves when a request takes
 * time or the code sleeps.
 */
function installAppsScript(
  responses: Array<FakeResponse | Error>,
  options: {
    /** The open tab, when there is one before the import. */
    active?: FakeTab;
    /** The selected cells' display values. */
    selection?: unknown[][];
    /** Where the clock starts, epoch ms. */
    now?: number;
  } = {},
): {
  urls: string[];
  bodies: Array<Record<string, unknown> | null>;
  /** Every setValues call: the row it started at and the rows it wrote. */
  written: Array<{ row: number; rows: unknown[][] }>;
  /** The first tab the import created, as it stands: header row first. */
  sheet: () => unknown[][];
  /** Every tab the import created. */
  inserted: FakeTab[];
  /** Every setNumberFormat call. */
  formats: Array<{ row: number; column: number; rows: number; columns: number; format: string }>;
  sleeps: number[];
  /** What had happened, in order: "fetch", "format", "write" and "flush". */
  events: string[];
} {
  const urls: string[] = [];
  const bodies: Array<Record<string, unknown> | null> = [];
  const written: Array<{ row: number; rows: unknown[][] }> = [];
  const formats: Array<{ row: number; column: number; rows: number; columns: number; format: string }> = [];
  const inserted: FakeTab[] = [];
  const sleeps: number[] = [];
  const events: string[] = [];
  const g = globalThis as unknown as Record<string, unknown>;
  let now = options.now ?? 1_000_000;
  Date.now = () => now;

  g.PropertiesService = { getUserProperties: () => ({ getProperty: () => "jbe_test_key" }) };
  g.Utilities = {
    sleep: (ms: number) => {
      sleeps.push(ms);
      now += ms;
    },
  };
  g.UrlFetchApp = {
    fetch: (url: string, options: { method?: string; payload?: string } = {}) => {
      urls.push(url);
      events.push("fetch");
      bodies.push(options.payload ? (JSON.parse(options.payload) as Record<string, unknown>) : null);
      const next = responses.shift();
      if (!next) throw new Error(`unexpected request ${url}`);
      if (next instanceof Error) throw next;
      now += next.takesMs ?? 0;
      return {
        getResponseCode: () => next.status,
        getContentText: () => (next.body === undefined ? "" : JSON.stringify(next.body)),
        getAllHeaders: () => next.headers ?? {},
      };
    },
  };
  const fakeSheet = (tab: FakeTab) => ({
    getName: () => tab.name,
    getLastRow: () => tab.grid.length,
    getLastColumn: () => Math.max(0, ...tab.grid.map((r) => r.length)),
    getRange: (row: number, column: number, numRows: number, numColumns: number) => {
      const cells = (): unknown[][] =>
        Array.from({ length: numRows }, (_, i) =>
          Array.from({ length: numColumns }, (_, j) => tab.grid[row - 1 + i]?.[column - 1 + j] ?? ""),
        );
      return {
        setValues: (rows: unknown[][]) => {
          assert.equal(rows.length, numRows, "setValues needs exactly the rows its range has");
          for (const r of rows) assert.equal(r.length, numColumns);
          events.push("write");
          written.push({ row, rows });
          rows.forEach((r, i) => {
            const target = (tab.grid[row - 1 + i] ??= []);
            r.forEach((value, j) => (target[column - 1 + j] = value));
          });
        },
        getValues: cells,
        getDisplayValues: () => cells().map((r) => r.map((value) => String(value ?? ""))),
        setFontWeight: () => undefined,
        setNumberFormat: (format: string) => {
          events.push("format");
          formats.push({ row, column, rows: numRows, columns: numColumns, format });
        },
      };
    },
    setFrozenRows: () => undefined,
    autoResizeColumns: () => undefined,
  });

  const tabs: FakeTab[] = options.active ? [options.active] : [];
  let active: FakeTab | null = options.active ?? null;
  g.SpreadsheetApp = {
    flush: () => void events.push("flush"),
    getActiveSheet: () => (active ? fakeSheet(active) : null),
    getActiveRange: () =>
      options.selection
        ? { getDisplayValues: () => (options.selection as unknown[][]).map((r) => r.map((v) => String(v ?? ""))) }
        : null,
    getActiveSpreadsheet: () => ({
      getSheetByName: (name: string) => tabs.find((tab) => tab.name === name) ?? null,
      insertSheet: (name: string) => {
        const tab: FakeTab = { name, grid: [] };
        tabs.push(tab);
        inserted.push(tab);
        // As in Sheets, a new tab becomes the open one.
        active = tab;
        return fakeSheet(tab);
      },
    }),
  };

  const sheet = (): unknown[][] => inserted[0]?.grid ?? [];

  return { urls, bodies, written, sheet, inserted, formats, sleeps, events };
}

const billedPage = (batch: Job[], totalPages: number): FakeResponse => ({
  status: 200,
  body: { jobs: batch, total_pages: totalPages },
  headers: { "X-Credits-Deducted": String(batch.length * 3), "X-Credits-Balance": "40" },
});

describe("importJobs", () => {
  it("writes the pages already billed when a later page returns 402", () => {
    const { urls, bodies, sheet } = installAppsScript([
      billedPage(jobs(100), 3),
      { status: 402, body: { detail: "Balance too low." } },
    ]);

    const result = importJobs({ limit: 300 });

    assert.equal(sheet().length, 101, "header row + the 100 fetched jobs");
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
    const { sheet } = installAppsScript([billedPage(jobs(51), 2), billedPage(jobs(51, 51), 2)]);

    importJobs({ limit: 101 });

    assert.equal(sheet().length, 102, "header + 101 jobs");
    assert.equal((sheet()[101] as unknown[])[13], "job-100", "the last job written is the 101st");
  });

  it("writes each page to the sheet before asking for the next", () => {
    const { written, sheet, events } = installAppsScript([
      billedPage(jobs(100), 3),
      billedPage([job({ id: "job-100", title: "=HYPERLINK(\"https://evil.test\")" }), ...jobs(99, 101)], 3),
      billedPage(jobs(100, 200), 3),
    ]);

    const result = importJobs({ limit: 300 });

    assert.equal(result.ok, true);
    assert.deepEqual(events, ["fetch", "write", "flush", "fetch", "write", "flush", "fetch", "write", "flush"]);
    assert.deepEqual(
      written.map(({ row, rows }) => [row, rows.length]),
      [
        [1, 101],
        [102, 100],
        [202, 100],
      ],
      "one batched setValues per page, the header only with the first",
    );
    const rows = sheet();
    assert.equal(rows.length, 301);
    assert.equal(rows[0][0], "Title");
    assert.equal(rows.filter((row) => row[0] === "Title").length, 1, "one header row");
    assert.equal(rows[101][0], `'=HYPERLINK("https://evil.test")`, "a later page is neutralised too");
    assert.deepEqual(
      rows.slice(1).map((row) => row[13]),
      Array.from({ length: 300 }, (_, n) => `job-${n}`),
    );
  });

  it("stops starting pages near Google's time limit and says so, keeping what it wrote", () => {
    // Each page takes 2.5 minutes: page 3 would start 5 minutes in, past the 4.5-minute budget.
    const slow = (batch: Job[]): FakeResponse => ({ ...billedPage(batch, 10), takesMs: 150_000 });
    const { urls, sheet } = installAppsScript([slow(jobs(100)), slow(jobs(100, 100)), slow(jobs(100, 200))]);

    const result = importJobs({ limit: 1000 });

    assert.equal(urls.length, 2);
    assert.equal(sheet().length, 201);
    assert.equal(result.ok, false);
    assert.match(result.message, /^Imported 200 of the 1000 jobs requested\. The import stopped early/);
    assert.match(result.message, /6-minute limit/);
    assert.match(result.message, /Cost: \$0\.60 \(600 credits\)/);
  });

  it("does not wait out a Retry-After that ends past the time budget", () => {
    const { urls, sleeps, sheet } = installAppsScript([
      { ...billedPage(jobs(100), 3), takesMs: 260_000 },
      { status: 429, headers: { "Retry-After": "30" } },
    ]);

    const result = importJobs({ limit: 300 });

    assert.equal(urls.length, 2);
    assert.deepEqual(sleeps, [], "the 30-second wait would end after the budget");
    assert.equal(sheet().length, 101);
    assert.match(result.message, /^Imported 100 of the 300 jobs requested before the import stopped\. Jobo rate limit/);
  });

  it("never replays a search after a 503 without Retry-After, a 500 or a dropped connection", () => {
    for (const failed of [
      { status: 503, body: { title: "Search took too long" } } as FakeResponse,
      { status: 503 } as FakeResponse,
      { status: 500, body: { title: "Internal error" } } as FakeResponse,
      { status: 502 } as FakeResponse,
      new Error("Address unavailable"),
    ]) {
      const { urls, sheet } = installAppsScript([billedPage(jobs(100), 2), failed]);

      const result = importJobs({ limit: 200 });

      assert.equal(urls.length, 2, `${JSON.stringify(failed)} is not retried: the page may already be billed`);
      assert.equal(sheet().length, 101);
      assert.equal(result.ok, false);
      assert.match(result.message, /^Imported 100 of the 200 jobs requested before the import stopped/);
    }
  });

  it("retries a search on a 429 and on a 503 that carries Retry-After, waiting as told", () => {
    const { urls, sleeps, sheet } = installAppsScript([
      { status: 503, headers: { "Retry-After": "5" }, body: { title: "Service unavailable" } },
      { status: 429, headers: { "Retry-After": "2" } },
      billedPage(jobs(10), 1),
    ]);

    const result = importJobs({ limit: 10 });

    assert.equal(urls.length, 3);
    assert.deepEqual(sleeps, [5000, 2000]);
    assert.equal(sheet().length, 11);
    assert.equal(result.ok, true);
  });

  it("gives up after three attempts", () => {
    const busy: FakeResponse = { status: 503, headers: { "Retry-After": "5" } };
    const { urls, sheet } = installAppsScript([billedPage(jobs(100), 2), busy, busy, busy]);

    const result = importJobs({ limit: 200 });

    assert.equal(urls.length, 4, "page 2 is tried three times");
    assert.equal(sheet().length, 101);
    assert.match(result.message, /^Imported 100 of the 200 jobs requested/);
    assert.match(result.message, /HTTP 503/);
  });

  it("writes the pages already billed when a later fetch throws", () => {
    const { sheet } = installAppsScript([billedPage(jobs(100), 2), new Error("Address unavailable")]);

    const result = importJobs({ limit: 200 });

    assert.equal(sheet().length, 101);
    assert.equal(result.ok, false);
    assert.match(result.message, /Jobo could not be reached: Address unavailable/);
  });

  it("writes nothing when the first page fails", () => {
    const { written } = installAppsScript([{ status: 402, body: { detail: "Balance too low." } }]);

    const result = importJobs({ limit: 300 });

    assert.equal(written.length, 0);
    assert.equal(result.ok, false);
    assert.match(result.message, /^Your Jobo wallet balance is too low/);
  });

  it("reports a complete import as before", () => {
    const { sheet } = installAppsScript([billedPage(jobs(100), 2), billedPage(jobs(50, 100), 2)]);

    const result = importJobs({ limit: 150 });

    assert.equal(sheet().length, 151);
    assert.equal(result.ok, true);
    assert.match(result.message, /^Imported 150 jobs\. Cost: \$0\.45 \(450 credits\)\./);
  });
});

function signal(overrides: Partial<HiringSignal> = {}): HiringSignal {
  return {
    id: "sig-1",
    type: "hiring_surge",
    detected_on: "2026-11-10",
    created_at: "2026-11-11T02:14:09.1234567Z",
    window_start: "2026-11-04",
    window_end: "2026-11-10",
    strength: 88,
    company: {
      id: "company-1",
      name: "Northwind Health",
      domain: "northwind.example",
      country_code: "US",
      size: "1-10",
      industry: "Hospitals and Health Care",
      industries: ["Hospitals and Health Care"],
      categories: [],
      is_agency: false,
      funding_stage: "Series B",
      founded_year: 2012,
    },
    details: { new_jobs: 17, baseline_weekly_avg: 3.1 },
    focus: {
      functions: ["technology", "sales"],
      title: "Senior Data Engineer",
      seniority: "senior",
      topics: ["snowflake"],
      job_countries: ["US", "CA"],
      job_regions: [],
      job_cities: [],
    },
    evidence_job_ids: ["job-a", "job-b"],
    ...overrides,
  };
}

function signalBatch(count: number, offset = 0): HiringSignal[] {
  return Array.from({ length: count }, (_, n) => signal({ id: `sig-${offset + n}` }));
}

const signalsPage = (batch: HiringSignal[], nextCursor?: string, headers: Record<string, string> = {}): FakeResponse => ({
  status: 200,
  body: { signals: batch, has_more: nextCursor !== undefined, ...(nextCursor ? { next_cursor: nextCursor } : {}) },
  headers,
});

/** The query of a signals request, as the API reads it. */
const queryOf = (url: string): Record<string, string> => Object.fromEntries(new URL(url).searchParams);

const column = (header: string): number => SIGNAL_COLUMNS.findIndex((c) => c.header === header);

/** A row of a signals tab the add-on wrote earlier. */
function tabRow(id: string, createdAt: string): string[] {
  return SIGNAL_COLUMNS.map((c) => (c.header === "Signal ID" ? id : c.header === "Created at" ? createdAt : "x"));
}

function signalsTab(rows: string[][], name = "Jobo Signals"): FakeTab {
  return { name, grid: [SIGNAL_COLUMNS.map((c) => c.header), ...rows] };
}

describe("hiring signals: companies", () => {
  it("cleans domains: trimmed, lower-cased, de-duplicated, without scheme, www., path or port", () => {
    const typed = "https://www.Stripe.com/careers?x=1\nstripe.com, shopify.com;  ACME.co.uk:443 \n\nhttp://user@sub.example.org/#a\tWebsite";

    assert.deepEqual(cleanDomains(typed), {
      domains: ["stripe.com", "shopify.com", "acme.co.uk", "sub.example.org"],
      rejected: ["Website"],
    });
    assert.deepEqual(cleanDomains(["notion.so", null, "", "www.notion.so/product", "localhost"]), {
      domains: ["notion.so"],
      rejected: ["localhost"],
    });
    assert.deepEqual(cleanDomains(undefined), { domains: [], rejected: [] });
  });

  it("fills the box from the selected cells, leaving out a header and blank cells", () => {
    installAppsScript([], {
      selection: [["Website", "Notes"], ["https://www.stripe.com/", ""], ["", ""], ["Stripe.com", "notion.so/product"]],
    });

    assert.deepEqual(getSelectedDomains(), { domains: ["stripe.com", "notion.so"], skipped: 2 });
  });

  it("finds no domains when nothing is selected", () => {
    installAppsScript([]);

    assert.deepEqual(getSelectedDomains(), { domains: [], skipped: 0 });
  });
});

describe("importSignals", () => {
  it("sends the sidebar's filters as the query of GET /api/signals", () => {
    const { urls, bodies } = installAppsScript([signalsPage([signal()])]);

    importSignals({
      domains: "stripe.com\nshopify.com",
      types: ["hiring_surge", "new_country"],
      countries: ["us", " de "],
      companySizes: ["51-200", "201-500"],
      functions: ["Technology"],
      minStrength: 70,
      detectedAfter: "2026-11-01",
      excludeAgencies: true,
      limit: 50,
    });

    assert.equal(urls.length, 1);
    assert.equal(new URL(urls[0]).origin + new URL(urls[0]).pathname, "https://connect.jobo.world/api/signals");
    assert.equal(bodies[0], null, "a GET");
    assert.deepEqual(queryOf(urls[0]), {
      types: "hiring_surge,new_country",
      domains: "stripe.com,shopify.com",
      countries: "US,DE",
      company_sizes: "51-200,201-500",
      functions: "Technology",
      exclude_agencies: "true",
      min_strength: "70",
      detected_after: "2026-11-01",
      limit: "50",
    });
  });

  it("sends nothing for an untouched field", () => {
    const { urls } = installAppsScript([signalsPage([])]);

    importSignals({
      domains: " \n ",
      types: ["leadership_hire"],
      countries: [],
      companySizes: [],
      functions: [],
      minStrength: 0,
      detectedAfter: "",
      excludeAgencies: false,
    });

    assert.deepEqual(queryOf(urls[0]), { types: "leadership_hire", limit: "50" });
  });

  it("reads more than 100 domains in requests of at most 100 that share How many signals", () => {
    const domains = Array.from({ length: 250 }, (_, n) => `company${n}.example`);
    const { urls, sheet } = installAppsScript([
      signalsPage(signalBatch(150), "c1"),
      signalsPage(signalBatch(10, 150)),
      signalsPage(signalBatch(100, 160), "c2"),
    ]);

    const result = importSignals({ domains: domains.join("\n"), limit: 260 });

    assert.equal(urls.length, 3, "the third chunk is never asked for: the budget is spent");
    const [first, second, third] = urls.map(queryOf);
    assert.deepEqual(first.domains.split(","), domains.slice(0, 100));
    assert.equal(first.limit, "200", "at most 200 a page");
    assert.equal(first.cursor, undefined);
    assert.deepEqual(second.domains.split(","), domains.slice(0, 100), "the same filters on the next page");
    assert.equal(second.cursor, "c1");
    assert.equal(second.limit, "110", "never more than is left of the 260");
    assert.deepEqual(third.domains.split(","), domains.slice(100, 200));
    assert.equal(third.cursor, undefined, "a new chunk starts from its first page");
    assert.equal(third.limit, "100");

    assert.equal(sheet().length, 261, "header + 260 signals");
    assert.equal(result.ok, true);
    assert.match(result.message, /^Imported 260 signals into “Jobo Signals”\. More signals matched than the 260 asked for\./);
    assert.match(result.message, /will not pick up the ones this import did not reach/);
  });

  it("says how many domains were never read when the budget runs out between chunks", () => {
    const domains = Array.from({ length: 150 }, (_, n) => `company${n}.example`);
    const { urls } = installAppsScript([signalsPage(signalBatch(30))]);

    const result = importSignals({ domains, limit: 30 });

    assert.equal(urls.length, 1);
    assert.match(result.message, /The 30 signals asked for were reached before 50 domains could be read\./);
  });

  it("writes one row per signal with every column, the summary from the core, and nothing that runs as a formula", () => {
    const { sheet, formats, events } = installAppsScript([
      signalsPage([signal({ company: { ...signal().company, name: '=HYPERLINK("https://evil.test")' } })]),
    ]);

    const result = importSignals({ types: ["hiring_surge"] });

    assert.equal(result.ok, true);
    const [header, row] = sheet();
    assert.deepEqual(header, [
      "Detected",
      "Company",
      "Domain",
      "Signal",
      "Summary",
      "Strength",
      "Strength band",
      "Country",
      "Company size",
      "Industry",
      "Funding stage",
      "Functions",
      "Role",
      "Seniority",
      "Topics",
      "Job countries",
      "Evidence job IDs",
      "Created at",
      "Signal ID",
    ]);
    assert.deepEqual(row, [
      "2026-11-10",
      `'=HYPERLINK("https://evil.test")`,
      "northwind.example",
      "Hiring surge",
      "17 new jobs this week vs 3.1 a week normally",
      88,
      "Very strong",
      "US",
      "1-10",
      "Hospitals and Health Care",
      "Series B",
      "technology, sales",
      "Senior Data Engineer",
      "senior",
      "snowflake",
      "US, CA",
      "job-a, job-b",
      "2026-11-11T02:14:09.1234567Z",
      "sig-1",
    ]);

    // Plain text before the values land, so Sheets keeps all seven decimal
    // places and does not read "1-10" as a date.
    assert.deepEqual(
      formats.map((f) => [f.row, f.column, f.rows, f.columns, f.format]),
      [
        [1, column("Company size") + 1, 2, 1, "@"],
        [1, column("Created at") + 1, 2, 1, "@"],
        [1, column("Signal ID") + 1, 2, 1, "@"],
      ],
    );
    assert.deepEqual(events, ["fetch", "format", "format", "format", "write", "flush"]);
  });

  it("never fails on a signal type it does not know", () => {
    const unknown = { ...signal({ id: "sig-new", type: "talent_exodus", details: { anything: 1 } }), focus: undefined };
    const { sheet } = installAppsScript([signalsPage([unknown as unknown as HiringSignal])]);

    const result = importSignals({ domains: "northwind.example" });

    assert.equal(result.ok, true);
    const row = sheet()[1];
    assert.equal(row[column("Signal")], "Talent exodus");
    assert.equal(row[column("Summary")], "Talent exodus");
    assert.equal(row[column("Functions")], "");
    assert.equal(row[column("Signal ID")], "sig-new");
  });

  it("writes each page before asking for the next, and keeps them when a later page fails", () => {
    const { urls, sheet, events } = installAppsScript([signalsPage(signalBatch(200), "c1"), { status: 500 }]);

    const result = importSignals({ types: ["hiring_surge"], limit: 400 });

    assert.equal(urls.length, 2, "a billed read is not replayed after a 500");
    assert.deepEqual(events.slice(0, 6), ["fetch", "format", "format", "format", "write", "flush"]);
    assert.equal(events[6], "fetch");
    assert.equal(sheet().length, 201);
    assert.equal(result.ok, false);
    assert.match(result.message, /^Imported 200 signals into “Jobo Signals”\. The import stopped early: Jobo returned HTTP 500\./);
  });

  it("stops starting requests near Google's time limit and says so", () => {
    const slow = (batch: HiringSignal[], cursor: string): FakeResponse => ({ ...signalsPage(batch, cursor), takesMs: 150_000 });
    const { urls, sheet } = installAppsScript([
      slow(signalBatch(200), "c1"),
      slow(signalBatch(200, 200), "c2"),
      slow(signalBatch(200, 400), "c3"),
    ]);

    const result = importSignals({ types: ["hiring_surge"], limit: 1000 });

    assert.equal(urls.length, 2);
    assert.equal(sheet().length, 401);
    assert.equal(result.ok, false);
    assert.match(result.message, /6-minute limit/);
  });

  it("reports what was billed and what is left of the included signals and the wallet", () => {
    installAppsScript([
      signalsPage(signalBatch(42), undefined, {
        "X-Credits-Deducted": "360",
        "X-Credits-Balance": "4200",
        "X-Quota-Limit": "250",
        "X-Quota-Remaining": "0",
      }),
    ]);

    const result = importSignals({ types: ["hiring_surge"] });

    assert.deepEqual(result, {
      ok: true,
      message:
        "Imported 42 signals into “Jobo Signals”. Billed for 42 signals. Cost: $0.36 (360 credits). " +
        "Wallet balance: $4.20. Included signals left this month: 0 of 250.",
    });
  });

  it("does not claim a charge the API did not report", () => {
    installAppsScript([signalsPage(signalBatch(3))]);

    const result = importSignals({ types: ["hiring_surge"] });

    assert.equal(result.message, "Imported 3 signals into “Jobo Signals”. Read 3 signals; Jobo reported no charge for them.");
  });

  it("creates no tab and bills nothing when nothing matches", () => {
    const { inserted } = installAppsScript([signalsPage([])]);

    const result = importSignals({ domains: "stripe.com" });

    assert.deepEqual(result, { ok: true, message: "No signals matched those filters. Nothing was billed." });
    assert.equal(inserted.length, 0);
  });

  it("names the next free tab when a Jobo Signals tab exists", () => {
    const { inserted } = installAppsScript([signalsPage([signal()])], { active: signalsTab([]) });

    const result = importSignals({ domains: "stripe.com" });

    assert.equal(inserted[0].name, "Jobo Signals 2");
    assert.match(result.message, /into “Jobo Signals 2”/);
  });

  it("requires domains or a filter that narrows the read, before anything is read", () => {
    const { urls } = installAppsScript([]);

    const result = importSignals({ minStrength: 80, excludeAgencies: true, detectedAfter: "2026-11-01" });

    assert.equal(result.ok, false);
    assert.match(result.message, /^Add the website domains of the companies you want, or choose at least one signal type/);
    assert.equal(urls.length, 0);
  });

  it("refuses a value that contains a comma, what is not a domain, and a country that is not a code", () => {
    const { urls, inserted } = installAppsScript([]);

    const comma = importSignals({ domains: "stripe.com", functions: ["Sales, Marketing"] });
    assert.equal(comma.ok, false);
    assert.match(comma.message, /"Sales, Marketing" cannot be used in Functions/);

    const notADomain = importSignals({ domains: "stripe.com\nAcme Corp" });
    assert.match(notADomain.message, /^Not a website domain: Acme, Corp\./);

    const country = importSignals({ domains: "stripe.com", countries: ["Germany"] });
    assert.match(country.message, /two-letter codes.*"GERMANY" is not one/);

    assert.equal(urls.length, 0);
    assert.equal(inserted.length, 0);
  });

  it("explains a 402, a key without an account, a sandbox key, a rate limit and a bad filter", () => {
    const cases: Array<[FakeResponse, RegExp]> = [
      [{ status: 402, body: { error: "Insufficient credits" } }, /included signals and your Jobo wallet do not cover.*nothing was charged/],
      [
        { status: 403, body: { title: "Account API key required", code: "account_key_required" } },
        /billed to a Jobo account, and this API key does not belong to one/,
      ],
      [{ status: 403, body: { title: "Forbidden" } }, /Jobo rejected the API key/],
      [{ status: 404, body: { title: "Not available in the sandbox", code: "not_available_in_sandbox" } }, /not available with a sandbox key/],
      [
        { status: 400, body: { title: "One or more validation errors occurred.", errors: { min_strength: ["min_strength must be between 1 and 100."] } } },
        /^min_strength must be between 1 and 100\.$/,
      ],
    ];
    for (const [response, expected] of cases) {
      const { urls, written } = installAppsScript([response]);

      const result = importSignals({ domains: "stripe.com" });

      assert.equal(result.ok, false);
      assert.match(result.message, expected);
      assert.equal(urls.length, 1, `${response.status} is not retried`);
      assert.equal(written.length, 0);
    }
  });

  it("waits out a rate limit before giving up with the limits that apply", () => {
    const limited: FakeResponse = { status: 429, headers: { "Retry-After": "1" } };
    const { urls, sleeps } = installAppsScript([limited, limited, limited]);

    const result = importSignals({ domains: "stripe.com" });

    assert.equal(urls.length, 3);
    assert.deepEqual(sleeps, [1000, 1000]);
    assert.match(result.message, /rate limit reached.*20 requests a minute, 200 an hour and 1,000 a day/);
  });

  it("describes failures directly too", () => {
    assert.match(describeSignalsFailure(402, null), /Top up your wallet/);
    assert.match(describeSignalsFailure(401, {}), /api-keys/);
    assert.equal(describeSignalsFailure(500, null), "Jobo returned HTTP 500.");
  });
});

describe("importSignals: only new", () => {
  const NOON = Date.parse("2026-11-20T12:00:00Z");

  it("reads the stored order from the tab's newest Created at, sent as written, and appends below the tab", () => {
    // The newest is two minutes old: there is no settling time to wait out
    // and no overlap, however fresh it is.
    const tab = signalsTab([
      tabRow("old-1", "2026-11-20T11:58:30.98765Z"),
      tabRow("old-2", "2026-11-20T11:58:30.987654Z"),
      tabRow("old-3", "2026-11-20T08:00:00Z"),
    ]);
    const { urls, written, inserted, formats } = installAppsScript(
      [signalsPage([signal({ id: "new-1", created_at: "2026-11-20T11:00:00.0000001Z" })])],
      { active: tab, now: NOON },
    );

    const result = importSignals({ domains: "stripe.com", onlyNew: true });

    assert.equal(queryOf(urls[0]).sort, "created_asc", "the order a read can be carried on in");
    assert.equal(queryOf(urls[0]).created_after, "2026-11-20T11:58:30.987654Z", "every digit, as written");
    assert.equal(inserted.length, 0, "no new tab");
    assert.deepEqual(
      written.map(({ row, rows }) => [row, rows.length]),
      [[5, 1]],
      "below the header and the three rows, with no second header",
    );
    assert.equal(tab.grid[4][column("Signal ID")], "new-1");
    assert.equal(tab.grid[4][column("Created at")], "2026-11-20T11:00:00.0000001Z");
    assert.ok(formats.every((f) => f.row === 5 && f.rows === 1));
    assert.equal(result.ok, true);
    assert.match(result.message, /^Added 1 new signal to “Jobo Signals”\./);
  });

  it("carries on exactly after an import that stopped at How many signals", () => {
    const tab = signalsTab([tabRow("old-1", "2026-11-20T11:00:00.000001Z")]);
    const { urls } = installAppsScript(
      [
        signalsPage(signalBatch(200), "c1"),
        signalsPage(signalBatch(50, 200), "c2"),
      ],
      { active: tab, now: NOON },
    );

    const result = importSignals({ domains: "stripe.com", onlyNew: true, limit: 250 });

    const [first, second] = urls.map(queryOf);
    assert.equal(second.cursor, "c1");
    assert.equal(second.sort, "created_asc", "the cursor belongs to the stored order, so every page names it");
    assert.equal(second.created_after, first.created_after, "the same bound on every page");
    assert.match(result.message, /^Added 250 new signals to “Jobo Signals”\. More signals matched than the 250 asked for\./);
    assert.match(result.message, /the next “Only new signals” import continues after the last one added/);
    assert.doesNotMatch(result.message, /will not pick up/);
  });

  it("still warns when it stops early with more than 100 domains, which are read one hundred at a time", () => {
    const domains = Array.from({ length: 150 }, (_, n) => `company${n}.example`);
    const tab = signalsTab([tabRow("old-1", "2026-11-20T11:00:00.000001Z")]);
    installAppsScript([signalsPage(signalBatch(100), "c1")], { active: tab, now: NOON });

    const result = importSignals({ domains: domains.join("\n"), onlyNew: true, limit: 100 });

    // The second hundred domains were never asked about, and their new
    // signals can be older than the newest one just written.
    assert.match(result.message, /will not pick up the ones this import did not reach/);
    assert.doesNotMatch(result.message, /continues after the last one added/);
  });

  it("leaves out a signal the tab already holds, and says it was billed again", () => {
    // Nothing in the tab is asked for again, so this should not happen. If the
    // API returns one anyway, it is not written twice and the message says so.
    const tab = signalsTab([tabRow("old-1", "2026-11-20T11:58:00.123456Z")]);
    const { urls } = installAppsScript(
      [
        signalsPage([signal({ id: "old-1" }), signal({ id: "new-1" }), signal({ id: "new-2" })], undefined, {
          "X-Quota-Limit": "250",
          "X-Quota-Remaining": "197",
          "X-Credits-Deducted": "0",
        }),
      ],
      { active: tab, now: NOON },
    );

    const result = importSignals({ domains: "stripe.com", onlyNew: true });

    assert.equal(queryOf(urls[0]).created_after, "2026-11-20T11:58:00.123456Z", "no reaching back");
    assert.deepEqual(
      tab.grid.slice(1).map((row) => row[column("Signal ID")]),
      ["old-1", "new-1", "new-2"],
      "old-1 is not written twice",
    );
    assert.equal(
      result.message,
      "Added 2 new signals to “Jobo Signals”. 1 signal was already in the tab and left out, and is billed again. " +
        "Billed for 3 signals. Included signals left this month: 197 of 250.",
    );
  });

  it("says so when everything read was already in the tab", () => {
    const tab = signalsTab([tabRow("old-1", "2026-11-20T11:58:00Z")]);
    installAppsScript([signalsPage([signal({ id: "old-1" })], undefined, { "X-Credits-Deducted": "30" })], {
      active: tab,
      now: NOON,
    });

    const result = importSignals({ domains: "stripe.com", onlyNew: true });

    assert.equal(result.ok, true);
    assert.match(result.message, /^No new signals\. 1 signal was already in the tab and left out/);
    assert.equal(tab.grid.length, 2);
  });

  it("reads a signals tab that holds no signals yet like a new one: no bound, newest detection day first", () => {
    const { urls } = installAppsScript([signalsPage([])], { active: signalsTab([]), now: NOON });

    const result = importSignals({ domains: "stripe.com", onlyNew: true });

    assert.equal(queryOf(urls[0]).created_after, undefined);
    // The stored order without a bound would start at the oldest signal Jobo holds.
    assert.equal(queryOf(urls[0]).sort, undefined);
    assert.equal(result.message, "No new signals since the last import into “Jobo Signals”. Nothing was billed.");
  });

  it("refuses when the open tab is not a signals tab, or has no Created at it can read", () => {
    for (const active of [
      { name: "Leads", grid: [["Domain"], ["stripe.com"]] },
      signalsTab([tabRow("old-1", "yesterday")]),
    ]) {
      const { urls, inserted } = installAppsScript([], { active, now: NOON });

      const result = importSignals({ domains: "stripe.com", onlyNew: true });

      assert.equal(result.ok, false);
      assert.match(result.message, /[Uu]ntick Only new signals to import into a new tab/);
      assert.equal(urls.length, 0);
      assert.equal(inserted.length, 0);
    }
  });

  it("tells the sidebar whether the open tab is a signals tab", () => {
    installAppsScript([], { active: signalsTab([tabRow("a", "2026-11-20T11:00:00Z"), tabRow("b", "2026-11-20T11:01:00Z")]) });
    assert.deepEqual(getSignalsTabStatus(), { isSignalsTab: true, name: "Jobo Signals", rows: 2 });

    installAppsScript([], { active: { name: "Leads", grid: [["Company", "Domain"]] } });
    assert.deepEqual(getSignalsTabStatus(), { isSignalsTab: false, name: "", rows: 0 });
  });

  it("uses the newest created_at as it is written, whatever its age or precision", () => {
    for (const newest of ["2026-11-20T11:50:00.837281Z", "2026-11-20T11:50:00.83728Z", "2026-11-20T11:50:00Z"]) {
      assert.equal(signalsReadBound(newest), newest);
    }
    assert.equal(signalsReadBound("  2026-11-20T11:50:00.5Z "), "2026-11-20T11:50:00.5Z");
    assert.equal(signalsReadBound(null), null);
    assert.equal(signalsReadBound("yesterday"), null);
  });

  it("compares created_at values to the seventh decimal place", () => {
    assert.equal(
      newestCreatedAt(["2026-11-20T09:15:30Z", "2026-11-20T09:15:30.0000001Z", "2026-11-20T09:15:29.9999999Z", "", "n/a"]),
      "2026-11-20T09:15:30.0000001Z",
    );
    assert.equal(newestCreatedAt(["2026-11-20T09:15:30.5+01:00", "2026-11-20T08:15:30.4Z"]), "2026-11-20T09:15:30.5+01:00");
    assert.equal(newestCreatedAt([]), null);
  });
});

describe("hiring signals: sidebar", () => {
  it("adds both imports to the Extensions menu", () => {
    const items: Array<[string, string]> = [];
    (globalThis as unknown as Record<string, unknown>).SpreadsheetApp = {
      getUi: () => ({
        createAddonMenu: () => {
          const menu = {
            addItem: (caption: string, fn: string) => (items.push([caption, fn]), menu),
            addToUi: () => undefined,
          };
          return menu;
        },
      }),
    };

    onOpen();

    assert.deepEqual(items, [
      ["Import jobs…", "showSidebar"],
      ["Import hiring signals…", "showSignalsSidebar"],
    ]);
    for (const [, fn] of items) assert.equal(typeof (globalThis as unknown as Record<string, unknown>)[fn], "function");
  });

  it("opens the one sidebar on the view the menu item names", () => {
    const shown: Array<{ file: string; view: unknown; title: string }> = [];
    const g = globalThis as unknown as Record<string, unknown>;
    g.HtmlService = {
      createTemplateFromFile: (file: string) => {
        const template: Record<string, unknown> = {};
        template.evaluate = () => {
          const page = { file, view: template.view, title: "" };
          shown.push(page);
          const output = { setTitle: (title: string) => ((page.title = title), output) };
          return output;
        };
        return template;
      },
    };
    g.SpreadsheetApp = { getUi: () => ({ showSidebar: () => undefined }) };

    showSidebar();
    showSignalsSidebar();

    assert.deepEqual(shown, [
      { file: "Sidebar", view: "jobs", title: "Career Site Jobs" },
      { file: "Sidebar", view: "signals", title: "Hiring Signals" },
    ]);
  });

  it("groups every signal type the way the dashboard does, with the core's labels", () => {
    const groups = signalTypeGroups();

    assert.deepEqual(
      groups.map((group) => group.label),
      ["Growth", "Expansion", "Leadership", "Tech and skills", "Slowing or struggling"],
    );
    const offered = groups.flatMap((group) => group.types.map((type) => type.value));
    assert.deepEqual([...offered].sort(), [...SIGNAL_TYPES].sort(), "each type once, none missing");
    for (const group of groups) {
      for (const type of group.types) assert.equal(type.label, SIGNAL_TYPE_LABELS[type.value]);
    }
    assert.deepEqual(SIGNAL_TYPE_GROUPS[0].types, ["hiring_surge", "function_surge", "hiring_resumed"]);
  });

  it("hands the sidebar the signal types and company sizes", () => {
    installSignIn([], { JOBO_API_KEY: LIVE_KEY });

    const settings = getSettings();

    assert.deepEqual(settings.signalTypeGroups, signalTypeGroups());
    assert.deepEqual(settings.companySizes, HIRING_SIGNAL_FILTERS.find((f) => f.key === "company_sizes")?.options);
  });

  it("sends only filters the signals API has", () => {
    const keys = HIRING_SIGNAL_FILTERS.map((filter) => filter.key as string);
    for (const key of SHEETS_SIGNAL_FILTER_KEYS) assert.ok(keys.includes(key), `"${key}" is not a signals filter`);
  });
});

describe("free requests", () => {
  it("retry a dropped connection and a 500, which cost nothing to replay", () => {
    installAppsScript([
      new Error("Address unavailable"),
      { status: 500 },
      { status: 200, body: { suggestions: [{ display_name: "Berlin, Germany", city: "Berlin", region: null, country: "Germany" }] } },
    ]);

    const suggest = (globalThis as unknown as Record<string, (q: string) => { suggestions: unknown[] }>).suggestLocations;
    assert.equal(suggest("berl").suggestions.length, 1);
  });

  it("do not retry a 503 the API sent without Retry-After", () => {
    const { urls } = installAppsScript([{ status: 503, body: { title: "Search took too long" } }]);

    const suggest = (globalThis as unknown as Record<string, (q: string) => { suggestions: unknown[] }>).suggestLocations;
    assert.deepEqual(suggest("berl").suggestions, []);
    assert.equal(urls.length, 1);
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

describe("retry policy (connector-core's)", () => {
  const failed = (status: number, body: unknown = null, retryAfterSeconds: number | null = null) => ({
    status,
    body,
    retryAfterSeconds,
  });

  it("never retries insufficient credits — the precheck would fail identically", () => {
    assert.equal(isRetryableResponse(failed(402), true), false);
  });

  it("never retries a void feed cursor", () => {
    assert.equal(isRetryableResponse(failed(409, { code: "feed_cursor_restart_required" }), true), false);
  });

  it("never retries an auth failure or a bad request", () => {
    for (const status of [400, 401, 403, 404]) assert.equal(isRetryableResponse(failed(status), true), false);
  });

  it("retries throttling and an outage that names a wait, billed or not", () => {
    for (const replaySafe of [true, false]) {
      assert.equal(isRetryableResponse(failed(429), replaySafe), true);
      assert.equal(isRetryableResponse(failed(503, { title: "Unavailable" }, 5), replaySafe), true);
    }
  });

  it("never retries the API's 503 without Retry-After: the search ran out of time", () => {
    for (const replaySafe of [true, false]) {
      assert.equal(isRetryableResponse(failed(503, { title: "Search took too long" }), replaySafe), false);
    }
  });

  it("retries a server fault only for a request that is free to replay", () => {
    assert.equal(isRetryableResponse(failed(500), true), true);
    assert.equal(isRetryableResponse(failed(502), true), true);
    assert.equal(isRetryableResponse(failed(500), false), false);
    assert.equal(isRetryableResponse(failed(502), false), false);
  });

  it("classifies a failed response as the core client would", () => {
    assert.ok(toCoreError(failed(402)) instanceof InsufficientCreditsError);
    assert.ok(toCoreError(failed(429)) instanceof RateLimitError);
    assert.ok(toCoreError(failed(409, { code: "feed_cursor_restart_required" })) instanceof FeedCursorRestartRequiredError);
    assert.equal(toCoreError(failed(503, null, 5)).retryAfterSeconds, 5);
  });
});

describe("built sidebar", () => {
  it("inlines the shared UI assets verbatim", () => {
    const dist = new URL("../dist/", import.meta.url);
    const ui = new URL("../../packages/connector-ui/", import.meta.url);
    const sidebar = readFileSync(new URL("Sidebar.html", dist), "utf8");

    assert.ok(sidebar.includes(readFileSync(new URL("combobox.css", ui), "utf8")));
    assert.ok(sidebar.includes(readFileSync(new URL("combobox.js", ui), "utf8").replace(/^export /gm, "")));
    assert.ok(!sidebar.includes("@inline:"), "every marker was replaced");
  });

  it("is a template with one scriptlet: the view, escaped, on the body", () => {
    const sidebar = readFileSync(new URL("../dist/Sidebar.html", import.meta.url), "utf8");

    // The template engine runs a scriptlet tag wherever it appears, inlined code included.
    assert.equal(sidebar.split("<?").length - 1, 1);
    assert.ok(sidebar.includes('<body data-view="<?= view ?>">'));
  });
});

describe("manifest", () => {
  it("describes the editor add-on the code implements", () => {
    const manifest = JSON.parse(readFileSync(new URL("../dist/appsscript.json", import.meta.url), "utf8")) as Record<
      string,
      unknown
    >;
    const globals = globalThis as unknown as Record<string, unknown>;

    // addOns (and its homepage trigger, which must return Cards) is for Google Workspace add-ons only.
    assert.equal(manifest.addOns, undefined);
    assert.equal(globals.onHomepage, undefined);
    assert.equal(typeof globals.onOpen, "function");
    assert.equal(typeof globals.onInstall, "function");
    assert.deepEqual(manifest.oauthScopes, [
      "https://www.googleapis.com/auth/spreadsheets.currentonly",
      "https://www.googleapis.com/auth/script.container.ui",
      "https://www.googleapis.com/auth/script.external_request",
    ]);
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
