/**
 * Jobo for Sheets — pulls filtered job data into the active spreadsheet.
 *
 * Apps Script constraints that shape this file:
 *
 *  - No npm at runtime. @jobo-ai/connector-core is bundled by build.mjs; this
 *    file uses its canonical filter values, types and key validation.
 *  - **The HTTP path here is synchronous on purpose.** connector-core's
 *    `JoboClient` is promise-based, and Apps Script has no event loop — a
 *    handler cannot return a promise across the `google.script.run` boundary,
 *    and there is no guarantee the microtask queue drains before the call
 *    returns. Rather than paper over that with a spin-wait, this file mirrors
 *    connector-core's error semantics against a blocking `UrlFetchApp` call.
 *    The shared behaviour that matters — which statuses are terminal — is
 *    reproduced explicitly below and covered by tests.
 */

import {
  EMPLOYMENT_TYPES,
  EXPERIENCE_LEVELS,
  WORK_MODELS,
  isValidApiKeyFormat,
  type Job,
} from "@jobo-ai/connector-core";

declare const SpreadsheetApp: GoogleAppsScript.Spreadsheet.SpreadsheetApp;
declare const HtmlService: GoogleAppsScript.HTML.HtmlService;
declare const UrlFetchApp: GoogleAppsScript.URL_Fetch.UrlFetchApp;
declare const PropertiesService: GoogleAppsScript.Properties.PropertiesService;
declare const Utilities: GoogleAppsScript.Utilities.Utilities;
declare const CacheService: GoogleAppsScript.Cache.CacheService;

const API_KEY_PROPERTY = "JOBO_API_KEY";
const BASE_URL = "https://connect.jobo.world";
const MAX_ATTEMPTS = 3;

/** Columns written to the sheet, in order. */
const COLUMNS: Array<{ header: string; value: (job: Job) => string | number }> = [
  { header: "Title", value: (j) => j.title },
  { header: "Company", value: (j) => j.company?.name ?? "" },
  { header: "Location", value: (j) => formatLocation(j) },
  { header: "Work model", value: (j) => j.workplace_type ?? "" },
  { header: "Employment type", value: (j) => j.employment_type ?? "" },
  { header: "Experience", value: (j) => j.experience_level ?? "" },
  { header: "Salary min", value: (j) => j.compensation?.min ?? "" },
  { header: "Salary max", value: (j) => j.compensation?.max ?? "" },
  { header: "Currency", value: (j) => j.compensation?.currency ?? "" },
  { header: "Posted", value: (j) => (j.date_posted ?? "").slice(0, 10) },
  { header: "Source", value: (j) => j.source ?? "" },
  { header: "Apply URL", value: (j) => j.apply_url ?? j.listing_url ?? "" },
  { header: "Company website", value: (j) => j.company?.website ?? "" },
  { header: "Job ID", value: (j) => j.id },
];

export function formatLocation(job: Job): string {
  const first = job.locations?.[0];
  if (!first) return job.workplace_type === "Remote" ? "Remote" : "";
  if (first.location) return first.location;
  return [first.city, first.region, first.country].filter(Boolean).join(", ");
}

export interface ApiResult {
  status: number;
  body: unknown;
  creditsDeducted: number | null;
  creditsBalance: number | null;
  quotaLimit: number | null;
  quotaRemaining: number | null;
  retryAfterSeconds: number | null;
}

/**
 * Which statuses are worth another attempt.
 *
 * Mirrors connector-core's `isRetryable`. The two deliberate exclusions are the
 * ones that look retryable and are not: 402, because the balance check prices
 * the requested page size so a retry fails identically, and 409
 * `feed_cursor_restart_required`, because a void cursor can never be replayed.
 */
export function isRetryableStatus(status: number, code: string | null): boolean {
  if (status === 402) return false;
  if (status === 409 && code === "feed_cursor_restart_required") return false;
  if (status === 401 || status === 403) return false;
  return status === 0 || status === 408 || status === 429 || status >= 500;
}

function readNumber(headers: Record<string, string>, name: string): number | null {
  const raw = headers[name];
  if (raw === undefined) return null;
  const parsed = Number(raw);
  return isFinite(parsed) ? parsed : null;
}

/** One blocking request, with connector-core's retry policy applied. */
function request(path: string, apiKey: string): ApiResult {
  let last: ApiResult | null = null;

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    // muteHttpExceptions is essential: without it Apps Script throws on any
    // non-2xx and the body — carrying the problem details and credit headers —
    // is lost, so nothing could be mapped to a useful message.
    const response = UrlFetchApp.fetch(`${BASE_URL}${path}`, {
      method: "get",
      headers: { "X-Api-Key": apiKey, Accept: "application/json" },
      muteHttpExceptions: true,
    });

    const rawHeaders = response.getAllHeaders() as Record<string, string | string[]>;
    const headers: Record<string, string> = {};
    for (const key of Object.keys(rawHeaders)) {
      const value = rawHeaders[key];
      headers[key.toLowerCase()] = Array.isArray(value) ? value.join(", ") : String(value);
    }

    let body: unknown = null;
    const text = response.getContentText();
    if (text) {
      try {
        body = JSON.parse(text);
      } catch {
        body = text;
      }
    }

    const status = response.getResponseCode();
    const code =
      body && typeof body === "object" && typeof (body as Record<string, unknown>).code === "string"
        ? ((body as Record<string, unknown>).code as string)
        : null;

    last = {
      status,
      body,
      creditsDeducted: readNumber(headers, "x-credits-deducted"),
      creditsBalance: readNumber(headers, "x-credits-balance"),
      quotaLimit: readNumber(headers, "x-quota-limit"),
      quotaRemaining: readNumber(headers, "x-quota-remaining"),
      retryAfterSeconds: readNumber(headers, "retry-after"),
    };

    if (status >= 200 && status < 300) return last;
    if (!isRetryableStatus(status, code) || attempt === MAX_ATTEMPTS) return last;

    // Honour Retry-After literally when present, otherwise exponential backoff.
    const waitMs = last.retryAfterSeconds != null ? last.retryAfterSeconds * 1000 : 1000 * Math.pow(2, attempt - 1);
    Utilities.sleep(Math.min(waitMs, 30000));
  }

  return last as ApiResult;
}

/** Human-readable message for a failed call. */
export function describeFailure(status: number, body: unknown): string {
  const detail =
    body && typeof body === "object"
      ? ((body as Record<string, unknown>).detail as string) ??
        ((body as Record<string, unknown>).error as string) ??
        null
      : null;

  if (status === 402) {
    return `Your Jobo wallet balance is too low for this import. ${detail ?? ""} Top up your wallet at enterprise.jobo.world, or import fewer jobs.`.trim();
  }
  if (status === 401 || status === 403) {
    return "Jobo rejected the API key. Check it is current at enterprise.jobo.world/api-keys.";
  }
  if (status === 429) {
    return "Jobo rate limit reached. Wait a minute and try again.";
  }
  return detail ?? `Jobo returned HTTP ${status}.`;
}

function apiKey(): string {
  const key = PropertiesService.getUserProperties().getProperty(API_KEY_PROPERTY) ?? "";
  if (!key) throw new Error("No Jobo API key saved. Add one in the sidebar.");
  return key;
}

/** Build a query string, dropping empty values. */
export function buildQuery(params: Record<string, string | number | undefined>): string {
  const parts: string[] = [];
  for (const key of Object.keys(params)) {
    const value = params[key];
    if (value === undefined || value === "") continue;
    parts.push(`${encodeURIComponent(key)}=${encodeURIComponent(String(value))}`);
  }
  return parts.length > 0 ? `?${parts.join("&")}` : "";
}

// ── UI entry points ──────────────────────────────────────────────────────────

function onOpen(): void {
  SpreadsheetApp.getUi().createAddonMenu().addItem("Import jobs…", "showSidebar").addToUi();
}

function onHomepage(): unknown {
  showSidebar();
  return null;
}

function showSidebar(): void {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar").setTitle("Career Site Jobs");
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── Called from the sidebar ──────────────────────────────────────────────────

function getSettings(): {
  hasKey: boolean;
  workModels: string[];
  employmentTypes: string[];
  experienceLevels: string[];
} {
  return {
    hasKey: Boolean(PropertiesService.getUserProperties().getProperty(API_KEY_PROPERTY)),
    workModels: WORK_MODELS.slice(),
    employmentTypes: EMPLOYMENT_TYPES.slice(),
    experienceLevels: EXPERIENCE_LEVELS.slice(),
  };
}

/**
 * Stored per-user, not per-document: spreadsheets get shared, and a document
 * property would hand the key to everyone it is shared with.
 */
function saveApiKey(key: string): { ok: boolean; message: string } {
  const trimmed = (key || "").trim();
  if (!isValidApiKeyFormat(trimmed)) {
    return {
      ok: false,
      message:
        "That does not look like a Jobo API key. Keys start with jbe_live_ or jbe_test_ and are 74 characters long.",
    };
  }
  PropertiesService.getUserProperties().setProperty(API_KEY_PROPERTY, trimmed);
  return { ok: true, message: "API key saved." };
}

function clearApiKey(): void {
  PropertiesService.getUserProperties().deleteProperty(API_KEY_PROPERTY);
}

const FILTER_OPTIONS_CACHE_KEY = "jobo_filter_options_v1";
/** 6 hours — the CacheService maximum. Source/industry churn is slower than this. */
const FILTER_OPTIONS_CACHE_SECONDS = 21600;

export interface FilterOptionsResult {
  ok: boolean;
  sources: string[];
  industries: string[];
}

function readOptionKeys(body: unknown, listField: string): string[] {
  if (!body || typeof body !== "object") return [];
  const list = (body as Record<string, unknown>)[listField];
  if (!Array.isArray(list)) return [];
  return list
    .map((entry) =>
      entry && typeof entry === "object" ? (entry as Record<string, unknown>).key : null,
    )
    .filter((key): key is string => typeof key === "string" && key.length > 0);
}

/**
 * Dynamic filter options for the sidebar comboboxes. Both lists are small
 * enough (≤250) to ship whole and filter client-side — a per-keystroke
 * google.script.run round trip is far too slow for typeahead. Cached in the
 * script cache (shared across users; the lists are not user-specific).
 * `ok: false` tells the sidebar to fall back to free-text chips.
 */
function getFilterOptions(): FilterOptionsResult {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(FILTER_OPTIONS_CACHE_KEY);
    if (cached) return JSON.parse(cached) as FilterOptionsResult;

    const key = apiKey();
    const filters = request("/api/connectors/filters", key);
    if (filters.status < 200 || filters.status >= 300) {
      return { ok: false, sources: [], industries: [] };
    }
    const industries = request("/api/connectors/filter-options/industries?limit=250", key);

    const result: FilterOptionsResult = {
      ok: true,
      sources: readOptionKeys(filters.body, "sources"),
      industries:
        industries.status >= 200 && industries.status < 300
          ? readOptionKeys(industries.body, "options")
          : [],
    };
    cache.put(FILTER_OPTIONS_CACHE_KEY, JSON.stringify(result), FILTER_OPTIONS_CACHE_SECONDS);
    return result;
  } catch {
    return { ok: false, sources: [], industries: [] };
  }
}

export interface LocationSuggestionResult {
  suggestions: Array<{
    display_name: string | null;
    city: string | null;
    region: string | null;
    country: string | null;
  }>;
}

/**
 * Location typeahead proxy. The only live suggest path in the sidebar —
 * Photon-normalized labels are worth the round trip. Never throws: an empty
 * list keeps the field usable as free text.
 */
function suggestLocations(q: string): LocationSuggestionResult {
  const trimmed = (q || "").trim();
  if (trimmed.length < 2) return { suggestions: [] };
  try {
    const key = apiKey();
    const result = request(
      "/api/connectors/locations/suggest" + buildQuery({ q: trimmed, limit: 5 }),
      key,
    );
    if (result.status < 200 || result.status >= 300) return { suggestions: [] };
    const body = result.body as { suggestions?: LocationSuggestionResult["suggestions"] };
    return { suggestions: Array.isArray(body.suggestions) ? body.suggestions : [] };
  } catch {
    return { suggestions: [] };
  }
}

/**
 * Search-surface filter keys this connector sends, matched against
 * connector-core's JOB_SEARCH_FILTERS by the drift-guard test. Deliberately
 * absent: posted_before / discovered_after / discovered_before — this is a
 * one-shot import UI, not an incremental sync, so backward-looking windows
 * add noise without a use case (posted_after alone covers "recent jobs").
 */
export const SHEETS_SUPPORTED_FILTER_KEYS: readonly string[] = [
  "q",
  "location",
  "sources",
  "work_model",
  "employment_type",
  "experience_level",
  "skills",
  "industries",
  "min_salary_usd",
  "max_salary_usd",
  "posted_after",
  "search_description",
];

interface ImportRequest {
  q?: string;
  location?: string;
  sources?: string[];
  skills?: string[];
  industries?: string[];
  workModels?: string[];
  employmentTypes?: string[];
  experienceLevels?: string[];
  minSalaryUsd?: number;
  maxSalaryUsd?: number;
  postedAfter?: string;
  searchDescriptions?: boolean;
  limit?: number;
}

function importJobs(req: ImportRequest): { ok: boolean; message: string } {
  try {
    const key = apiKey();
    const limit = Math.min(Math.max(req.limit ?? 100, 1), 1000);

    const jobs: Job[] = [];
    let page = 1;
    let creditsSpent = 0;
    let balance: number | null = null;
    let quotaLimit: number | null = null;
    let quotaRemaining: number | null = null;

    while (jobs.length < limit) {
      const path =
        "/api/jobs" +
        buildQuery({
          q: req.q,
          location: req.location,
          sources: (req.sources ?? []).join(","),
          skills: (req.skills ?? []).join(","),
          industries: (req.industries ?? []).join(","),
          work_model: (req.workModels ?? []).join(","),
          employment_type: (req.employmentTypes ?? []).join(","),
          experience_level: (req.experienceLevels ?? []).join(","),
          min_salary_usd: req.minSalaryUsd && req.minSalaryUsd > 0 ? req.minSalaryUsd : undefined,
          max_salary_usd: req.maxSalaryUsd && req.maxSalaryUsd > 0 ? req.maxSalaryUsd : undefined,
          posted_after: req.postedAfter || undefined,
          // true is the API default; only the opt-out travels on the wire.
          search_description: req.searchDescriptions === false ? "false" : undefined,
          page,
          page_size: Math.min(100, limit - jobs.length),
        });

      const result = request(path, key);
      if (result.status < 200 || result.status >= 300) {
        return { ok: false, message: describeFailure(result.status, result.body) };
      }

      creditsSpent += result.creditsDeducted ?? 0;
      balance = result.creditsBalance ?? balance;
      quotaLimit = result.quotaLimit ?? quotaLimit;
      quotaRemaining = result.quotaRemaining ?? quotaRemaining;

      const data = result.body as { jobs?: Job[]; total_pages?: number };
      const batch = data.jobs ?? [];
      for (const job of batch) jobs.push(job);

      if (batch.length === 0 || page >= (data.total_pages ?? 1)) break;
      page++;
    }

    if (jobs.length === 0) return { ok: true, message: "No jobs matched those filters." };

    writeSheet(jobs);

    const cost = creditsSpent > 0 ? ` Cost: $${(creditsSpent / 1000).toFixed(2)} (${creditsSpent} credits).` : "";
    const left = balance != null ? ` Wallet balance: $${(balance / 1000).toFixed(2)}.` : "";
    const included = quotaRemaining != null
      ? ` Included jobs: ${quotaRemaining}${quotaLimit != null ? ` of ${quotaLimit}` : ""} remaining.`
      : "";
    return { ok: true, message: `Imported ${jobs.length} jobs.${cost}${left}${included}` };
  } catch (err) {
    return { ok: false, message: (err as Error).message };
  }
}

function writeSheet(jobs: Job[]): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.insertSheet(uniqueSheetName(spreadsheet, "Jobo Jobs"));

  const rows: Array<Array<string | number>> = [COLUMNS.map((c) => c.header)];
  for (const job of jobs) rows.push(COLUMNS.map((c) => c.value(job)));

  sheet.getRange(1, 1, rows.length, COLUMNS.length).setValues(rows);
  sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, COLUMNS.length);
}

export function uniqueSheetName(
  spreadsheet: { getSheetByName: (name: string) => unknown },
  base: string,
): string {
  if (!spreadsheet.getSheetByName(base)) return base;
  for (let i = 2; i < 200; i++) {
    const candidate = `${base} ${i}`;
    if (!spreadsheet.getSheetByName(candidate)) return candidate;
  }
  return `${base} overflow`;
}

// Apps Script resolves handlers as globals, which a bundled module scope hides.
const globals = globalThis as unknown as Record<string, unknown>;
globals.onOpen = onOpen;
globals.onHomepage = onHomepage;
globals.showSidebar = showSidebar;
globals.getSettings = getSettings;
globals.saveApiKey = saveApiKey;
globals.clearApiKey = clearApiKey;
globals.getFilterOptions = getFilterOptions;
globals.suggestLocations = suggestLocations;
globals.importJobs = importJobs;
