import assert from "node:assert/strict";
import { describe, it } from "node:test";

import {
  SHEETS_SUPPORTED_FILTER_KEYS,
  buildQuery,
  companyPickerItem,
  describeFailure,
  flattenOccupations,
  formatLocation,
  importJobs,
  importPageSize,
  isRetryableStatus,
  neutralizeFormula,
  occupationLevelOptions,
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
