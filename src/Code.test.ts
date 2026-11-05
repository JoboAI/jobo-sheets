import assert from "node:assert/strict";
import { describe, it } from "node:test";

import {
  SHEETS_SUPPORTED_FILTER_KEYS,
  buildQuery,
  describeFailure,
  formatLocation,
  importJobs,
  isRetryableStatus,
  neutralizeFormula,
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
function installAppsScript(responses: Array<FakeResponse | Error>): { urls: string[]; written: unknown[][][] } {
  const urls: string[] = [];
  const written: unknown[][][] = [];
  const g = globalThis as unknown as Record<string, unknown>;

  g.PropertiesService = { getUserProperties: () => ({ getProperty: () => "jbe_test_key" }) };
  g.Utilities = { sleep: () => undefined };
  g.UrlFetchApp = {
    fetch: (url: string) => {
      urls.push(url);
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

  return { urls, written };
}

const billedPage = (batch: Job[], totalPages: number): FakeResponse => ({
  status: 200,
  body: { jobs: batch, total_pages: totalPages },
  headers: { "X-Credits-Deducted": String(batch.length * 3), "X-Credits-Balance": "40" },
});

describe("importJobs", () => {
  it("writes the pages already billed when a later page returns 402", () => {
    const { urls, written } = installAppsScript([
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
    assert.match(urls[1], /[?&]page=2&/);
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
    // backward-looking windows stay out deliberately (see Code.ts).
    const knownGaps = new Set(["posted_before", "discovered_after", "discovered_before"]);

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
