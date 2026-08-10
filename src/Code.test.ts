import assert from "node:assert/strict";
import { describe, it } from "node:test";

import {
  SHEETS_SUPPORTED_FILTER_KEYS,
  buildQuery,
  describeFailure,
  formatLocation,
  isRetryableStatus,
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
