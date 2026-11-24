/**
 * Jobo for Sheets — pulls filtered job data, and the hiring signals of the
 * companies a user cares about, into the active spreadsheet.
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
 *    Whether a failure is retried, and how long to wait, is still decided by
 *    connector-core (`isRetryable`, `nextDelayMs`), so the policy cannot drift.
 *  - **A run ends after 6 minutes.** An import writes each page as it arrives
 *    and stops starting new work after `IMPORT_TIME_BUDGET_MS`.
 */

import {
  COMPANY_CATEGORIES,
  COMPANY_SIZES,
  EMPLOYMENT_TYPES,
  EXPERIENCE_LEVELS,
  FeedCursorRestartRequiredError,
  InsufficientCreditsError,
  JoboApiError,
  JoboTransportError,
  LANGUAGES,
  NARROWING_SIGNAL_FILTER_KEYS,
  RateLimitError,
  SIGNAL_DEFAULT_LIMIT,
  SIGNAL_MAX_LIMIT,
  SIGNAL_MAX_LIST_VALUES,
  SIGNAL_POLL_SORT,
  SIGNAL_TYPES,
  WORK_MODELS,
  defaultRetryOptions,
  describeSignal,
  isRetryable,
  isValidApiKeyFormat,
  nextDelayMs,
  signalStrengthBand,
  signalTypeLabel,
  toJobSearchBody,
  toSignalsQuery,
  type CompanySuggestion,
  type HiringSignal,
  type HiringSignalFilters,
  type HiringSignalListResponse,
  type Job,
  type JobFacet,
  type JobSearchFilters,
  type JobSearchResponse,
  type OccupationTaxonomyResponse,
  type SearchSortValue,
} from "@jobo-ai/connector-core";

declare const SpreadsheetApp: GoogleAppsScript.Spreadsheet.SpreadsheetApp;
declare const HtmlService: GoogleAppsScript.HTML.HtmlService;
declare const UrlFetchApp: GoogleAppsScript.URL_Fetch.UrlFetchApp;
declare const PropertiesService: GoogleAppsScript.Properties.PropertiesService;
declare const Utilities: GoogleAppsScript.Utilities.Utilities;
declare const CacheService: GoogleAppsScript.Cache.CacheService;
declare const ScriptApp: GoogleAppsScript.Script.ScriptApp;

const API_KEY_PROPERTY = "JOBO_API_KEY";
/** The signed-in account's email, kept only to say who is connected. Absent for a pasted key. */
const ACCOUNT_EMAIL_PROPERTY = "JOBO_ACCOUNT_EMAIL";
const BASE_URL = "https://connect.jobo.world";
const MAX_ATTEMPTS = 3;

/**
 * One column of an import. `text` makes the column plain text before anything
 * is written to it, so Sheets keeps the value as sent instead of reading it as
 * a date or a number.
 */
interface Column<T> {
  header: string;
  value: (item: T) => string | number;
  text?: boolean;
}

/** Columns written to the sheet, in order. */
const COLUMNS: Array<Column<Job>> = [
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

/**
 * setValues() runs any string starting with "=" as a formula, and job titles are
 * scraped from pages anyone can post on (=HYPERLINK, =IMPORTDATA…). A leading
 * apostrophe makes Sheets keep the cell as text; numbers pass through untouched.
 */
export function neutralizeFormula(value: string | number): string | number {
  return typeof value === "string" && /^[=+\-@\t\r]/.test(value) ? `'${value}` : value;
}

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
 * A failed response as the error connector-core's client would throw for it,
 * so the core's own `isRetryable` and `nextDelayMs` decide what happens next.
 * The class matters: 402 and a void feed cursor are never retried, and a 429
 * always is.
 */
export function toCoreError(result: Pick<ApiResult, "status" | "body" | "retryAfterSeconds">): JoboApiError {
  const body = result.body;
  const code =
    body && typeof body === "object" && typeof (body as Record<string, unknown>).code === "string"
      ? ((body as Record<string, unknown>).code as string)
      : null;
  const base = {
    status: result.status,
    message: `Jobo returned HTTP ${result.status}.`,
    code,
    body,
    retryAfterSeconds: result.retryAfterSeconds,
  };
  if (result.status === 402) return new InsufficientCreditsError(base);
  if (result.status === 409 && code === "feed_cursor_restart_required") return new FeedCursorRestartRequiredError(base);
  if (result.status === 429) return new RateLimitError(base);
  return new JoboApiError(base);
}

/**
 * Whether a failed response is worth another attempt: connector-core's policy.
 *
 * `replaySafe: false` is a request billed for what it returns (the search).
 * The API charges before it answers, so a lost answer may already have been
 * paid for; such a request is replayed only on a 429 or a 503 that carries
 * `Retry-After`. A 503 without it is a search that ran out of time, and the
 * same search runs out of time again, so it is never retried either way.
 */
export function isRetryableResponse(
  result: Pick<ApiResult, "status" | "body" | "retryAfterSeconds">,
  replaySafe: boolean,
): boolean {
  return isRetryable(toCoreError(result), { retryRateLimit: true, replaySafe });
}

function readNumber(headers: Record<string, string>, name: string): number | null {
  const raw = headers[name];
  if (raw === undefined) return null;
  const parsed = Number(raw);
  return isFinite(parsed) ? parsed : null;
}

/** connector-core's backoff settings, with the longest single wait kept short for a 6-minute run. */
const RETRY_TIMING = {
  ...defaultRetryOptions,
  maxAttempts: MAX_ATTEMPTS,
  maxDelayMs: 30_000,
  // nextDelayMs never sleeps; the add-on waits with Utilities.sleep below.
  sleep: () => Promise.resolve(),
};

export interface RequestOptions {
  /**
   * False for a request billed for what it returns: the job search and the
   * signals read. Default true: every other endpoint this add-on calls is free.
   */
  replaySafe?: boolean;
  /** Epoch ms. No retry is started whose wait would end after it. */
  deadline?: number;
}

/** One HTTP attempt. Throws only when there is no response at all (UrlFetchApp's network faults). */
function fetchOnce(path: string, apiKey: string, json: unknown): ApiResult {
  // muteHttpExceptions is essential: without it Apps Script throws on any
  // non-2xx and the body — carrying the problem details and credit headers —
  // is lost, so nothing could be mapped to a useful message.
  const response = UrlFetchApp.fetch(`${BASE_URL}${path}`, {
    method: json === undefined ? "get" : "post",
    // X-Jobo-Client is how the API attributes this traffic to the add-on;
    // Apps Script owns the User-Agent and will not let us set one.
    headers: { "X-Api-Key": apiKey, Accept: "application/json", "X-Jobo-Client": "sheets" },
    ...(json === undefined ? {} : { contentType: "application/json", payload: JSON.stringify(json) }),
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

  // Like the core client: the header first, then a retry_after_seconds body field.
  const bodyRetryAfter =
    body && typeof body === "object" && typeof (body as Record<string, unknown>).retry_after_seconds === "number"
      ? ((body as Record<string, unknown>).retry_after_seconds as number)
      : null;

  return {
    status: response.getResponseCode(),
    body,
    creditsDeducted: readNumber(headers, "x-credits-deducted"),
    creditsBalance: readNumber(headers, "x-credits-balance"),
    quotaLimit: readNumber(headers, "x-quota-limit"),
    quotaRemaining: readNumber(headers, "x-quota-remaining"),
    retryAfterSeconds: readNumber(headers, "retry-after") ?? bodyRetryAfter,
  };
}

/**
 * One blocking request, with connector-core's retry policy applied. A `json`
 * body makes it a JSON POST (the search endpoint); otherwise a GET.
 *
 * Returns the last response, failed or not. Throws when the last attempt got
 * no response at all, and at once for a request that is not replay safe: the
 * charge may have been made before the connection dropped.
 */
function request(path: string, apiKey: string, json?: unknown, options: RequestOptions = {}): ApiResult {
  const replaySafe = options.replaySafe !== false;

  for (let attempt = 1; ; attempt++) {
    let result: ApiResult | null = null;
    let error: JoboApiError | JoboTransportError;
    try {
      result = fetchOnce(path, apiKey, json);
      if (result.status >= 200 && result.status < 300) return result;
      error = toCoreError(result);
    } catch (thrown) {
      const reason = thrown instanceof Error ? thrown.message : String(thrown);
      error = new JoboTransportError(`Jobo could not be reached: ${reason}`, thrown);
    }

    const retry = attempt < MAX_ATTEMPTS && isRetryable(error, { retryRateLimit: true, replaySafe });
    const waitMs = retry ? nextDelayMs(error, attempt, RETRY_TIMING) : 0;
    const inTime = options.deadline == null || Date.now() + waitMs < options.deadline;
    if (!retry || !inTime) {
      if (result) return result;
      throw error;
    }
    Utilities.sleep(waitMs);
  }
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
    return "Jobo rejected the API key. Click Disconnect at the bottom of the sidebar, then sign in again or paste a current key from https://enterprise.jobo.world/api-keys?src=sheets.";
  }
  if (status === 429) {
    return "Jobo rate limit reached. Wait a minute and try again.";
  }
  return detail ?? `Jobo returned HTTP ${status}.`;
}

function apiKey(): string {
  const key = PropertiesService.getUserProperties().getProperty(API_KEY_PROPERTY) ?? "";
  if (!key) throw new Error("Not connected to Jobo. Sign in or add an API key in the sidebar.");
  return key;
}

/** Build a query string, dropping empty values. */
export function buildQuery(params: Record<string, string | number | boolean | undefined>): string {
  const parts: string[] = [];
  for (const key of Object.keys(params)) {
    const value = params[key];
    if (value === undefined || value === "") continue;
    parts.push(`${encodeURIComponent(key)}=${encodeURIComponent(String(value))}`);
  }
  return parts.length > 0 ? `?${parts.join("&")}` : "";
}

// ── UI entry points ──────────────────────────────────────────────────────────
//
// This is an editor add-on: a menu under Extensions and an HTML sidebar. The
// manifest therefore has no `addOns` section; that section, and its homepage
// trigger that must return Cards, belong to Google Workspace add-ons.
//
// onOpen and onInstall can run before the user has authorised the add-on in
// this file (AuthMode.NONE), so they only build the menu.

export function onOpen(): void {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem("Import jobs…", "showSidebar")
    .addItem("Import hiring signals…", "showSignalsSidebar")
    .addToUi();
}

/** Installing from the Marketplace does not run onOpen for the file already open. */
function onInstall(): void {
  onOpen();
}

/**
 * Both menu items open the one sidebar, on a different view. The views share
 * the connection (sign-in, key, Disconnect), so it lives in one place.
 *
 * Sidebar.html is a template for this alone: the view is printed into
 * `data-view` on its body, where a printing scriptlet escapes it. A sidebar
 * cannot be told which menu item opened it any other way.
 */
function openSidebar(view: "jobs" | "signals"): void {
  const template = HtmlService.createTemplateFromFile("Sidebar");
  template.view = view;
  const html = template.evaluate().setTitle(view === "signals" ? "Hiring Signals" : "Career Site Jobs");
  SpreadsheetApp.getUi().showSidebar(html);
}

export function showSidebar(): void {
  openSidebar("jobs");
}

export function showSignalsSidebar(): void {
  openSidebar("signals");
}

// ── Called from the sidebar ──────────────────────────────────────────────────

export function getSettings(): {
  hasKey: boolean;
  /** The signed-in account, when the key came from Sign in with Jobo. */
  accountEmail: string | null;
  workModels: string[];
  employmentTypes: string[];
  experienceLevels: string[];
  companyCategories: string[];
  languages: Array<{ code: string; name: string }>;
  signalTypeGroups: SignalTypeGroup[];
  companySizes: string[];
} {
  const connection = getConnectionStatus();
  return {
    hasKey: connection.connected,
    accountEmail: connection.email,
    workModels: WORK_MODELS.slice(),
    employmentTypes: EMPLOYMENT_TYPES.slice(),
    experienceLevels: EXPERIENCE_LEVELS.slice(),
    companyCategories: COMPANY_CATEGORIES.slice(),
    languages: LANGUAGES.map((language) => ({ code: language.code, name: language.name })),
    signalTypeGroups: signalTypeGroups(),
    companySizes: COMPANY_SIZES.slice(),
  };
}

/**
 * Stored per-user, not per-document: spreadsheets get shared, and a document
 * property would hand the key to everyone it is shared with.
 *
 * The one way a key gets stored: the sidebar's pasted key and the key that
 * Sign in with Jobo returns both come through here. A newly saved key drops
 * the previous account's email and the app it was issued to; authCallback()
 * sets them again after a sign-in.
 */
export function saveApiKey(key: string): { ok: boolean; message: string } {
  const trimmed = (key || "").trim();
  if (!isValidApiKeyFormat(trimmed)) {
    return {
      ok: false,
      message:
        "That does not look like a Jobo API key. Keys start with jbe_live_ or jbe_test_ and are 74 characters long.",
    };
  }
  const properties = PropertiesService.getUserProperties();
  properties.setProperty(API_KEY_PROPERTY, trimmed);
  properties.deleteProperty(ACCOUNT_EMAIL_PROPERTY);
  properties.deleteProperty(OAUTH_CLIENT_PROPERTY);
  return { ok: true, message: "API key saved." };
}

/** Whether a key is stored, and whose account it is when the user signed in. */
export function getConnectionStatus(): { connected: boolean; email: string | null } {
  const properties = PropertiesService.getUserProperties();
  const connected = Boolean(properties.getProperty(API_KEY_PROPERTY));
  return { connected, email: connected ? properties.getProperty(ACCOUNT_EMAIL_PROPERTY) || null : null };
}

/**
 * Disconnects the add-on. A key that Sign in with Jobo stored is handed back to
 * Jobo first, which revokes it and removes the add-on from Developers → API
 * Keys in the dashboard. A pasted key is the user's own: it is forgotten here
 * and left alone at Jobo. The add-on is disconnected either way; `revoked` is
 * false only when a signed-in key could not be handed back, and the message
 * then says where to finish.
 */
export function disconnect(): { revoked: boolean; message: string } {
  const properties = PropertiesService.getUserProperties();
  const key = properties.getProperty(API_KEY_PROPERTY);
  const clientId = properties.getProperty(OAUTH_CLIENT_PROPERTY);
  const signedIn = Boolean(key && clientId);
  const revoked = !signedIn || revokeKey(key as string, clientId as string);

  properties.deleteProperty(API_KEY_PROPERTY);
  properties.deleteProperty(ACCOUNT_EMAIL_PROPERTY);
  properties.deleteProperty(OAUTH_CLIENT_PROPERTY);

  if (!signedIn) return { revoked, message: "" };
  return {
    revoked,
    message: revoked
      ? "Disconnected. The key this add-on held has been revoked."
      : "Disconnected here, but Jobo could not be reached to revoke the key. To finish, disconnect the add-on under Developers → API Keys in the Jobo dashboard.",
  };
}

/** Hands a signed-in key back (RFC 7009). Jobo answers 200 whether or not the key was still live. */
function revokeKey(key: string, clientId: string): boolean {
  // UrlFetchApp throws on a network fault even with muteHttpExceptions.
  try {
    const response = UrlFetchApp.fetch(OAUTH_REVOKE_URL, {
      method: "post",
      contentType: "application/x-www-form-urlencoded",
      headers: { "X-Jobo-Client": "sheets" },
      // Jobo only revokes a key for the app it was issued to.
      payload: buildQuery({ token: key, client_id: clientId }).slice(1),
      muteHttpExceptions: true,
    });
    return response.getResponseCode() === 200;
  } catch {
    return false;
  }
}

// ── Sign in with Jobo ────────────────────────────────────────────────────────
//
// OAuth 2.0 authorization code + PKCE (S256), as a public client: there is no
// client secret, and the access token that comes back *is* a Jobo API key. It
// is saved through saveApiKey(), so everything after sign-in is the pasted-key
// path. The key does not expire and is never refreshed.
//
// What Apps Script dictates:
//  - The redirect lands on /usercallback, which Apps Script serves for every
//    script project. The `state` token it issues names the function to run, and
//    Apps Script rejects a forged or expired token before that function is called.
//  - A sidebar cannot navigate the top window, so the sidebar shows the
//    authorization URL as a link that opens a new tab, then polls
//    getConnectionStatus() until authCallback() has stored the key.

const OAUTH_AUTHORIZE_URL = "https://enterprise.jobo.world/connect/authorize";
const OAUTH_TOKEN_URL = "https://enterprise.jobo.world/connect/token";
/** Says whether a client has a given redirect address registered. Public, no credentials. */
const OAUTH_CLIENT_LOOKUP_URL = "https://enterprise.jobo.world/api/auth/oauth/client";
const OAUTH_REGISTER_URL = "https://enterprise.jobo.world/connect/register";
const OAUTH_REVOKE_URL = "https://enterprise.jobo.world/connect/revoke";
/** The client Jobo registered for the published add-on. */
const OAUTH_CLIENT_ID = "google-sheets";
/** The client a copy of the add-on registered for itself: JSON `{ clientId, redirectUri }`, per script project. */
const SELF_REGISTERED_CLIENT_PROPERTY = "JOBO_OAUTH_SELF_CLIENT";
/**
 * User property: the app the stored key was issued to by a sign-in. Absent for
 * a pasted key. It is what lets disconnect() hand the key back.
 */
const OAUTH_CLIENT_PROPERTY = "JOBO_OAUTH_CLIENT";
const SELF_REGISTERED_CLIENT_NAME = "Google Sheets (Career Site Jobs)";
const RESOLVED_CLIENT_CACHE_KEY = "jobo_oauth_client";
const SIGN_IN_UNAVAILABLE = "Jobo could not be reached to start signing in. Try again, or paste an API key.";
const OAUTH_SCOPE = "api email";
/** The name authCallback is registered under on globalThis (see the end of this file). */
const OAUTH_CALLBACK_FUNCTION = "authCallback";
/**
 * How long one sign-in link works: the state token's documented maximum. The
 * sidebar asks for a fresh link well inside it, so a link is never near expiry
 * when it is clicked.
 */
const SIGN_IN_TIMEOUT_SECONDS = 3600;
const VERIFIER_CACHE_PREFIX = "jobo_oauth_verifier_";
const CALLBACK_PAGE_TITLE = "Career Site Jobs – Jobo";
const TRY_AGAIN_HINT = "Close this tab, then choose Sign in with Jobo in the sidebar to try again.";

/**
 * The redirect URI the Jobo server must have registered for `google-sheets`.
 * It is compared exactly, so it is always built from this script's own id.
 */
export function oauthRedirectUri(): string {
  return `https://script.google.com/macros/d/${ScriptApp.getScriptId()}/usercallback`;
}

/**
 * A PKCE code verifier: 96 hex characters from three random UUIDs, inside the
 * 43 to 128 unreserved characters RFC 7636 allows, with 366 random bits. Apps
 * Script has no crypto.getRandomValues; Utilities.getUuid() is documented as
 * java.util.UUID.randomUUID(), which draws from a cryptographically strong
 * generator.
 */
export function createCodeVerifier(): string {
  return [Utilities.getUuid(), Utilities.getUuid(), Utilities.getUuid()].join("").replace(/-/g, "");
}

/** The S256 challenge: BASE64URL(SHA-256(verifier)) without padding. */
export function codeChallenge(verifier: string): string {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, verifier, Utilities.Charset.US_ASCII);
  return Utilities.base64EncodeWebSafe(digest).replace(/=+$/, "");
}

/**
 * Which OAuth client this copy of the add-on signs in as.
 *
 * The redirect address contains the script id, and Jobo compares it exactly.
 * The published add-on is one script project, whose address Jobo has
 * registered for `google-sheets`. A copy installed from the source is a
 * different script project with an address Jobo cannot know in advance, so it
 * registers itself once (RFC 7591) and signs in as that client; the consent
 * screen then shows it as an unverified app, which is what it is.
 *
 * The registration is kept in script properties: one per script project,
 * whoever the user is. Signing in again as the same client returns the same key.
 */
export function resolveOAuthClientId(redirectUri: string): string {
  // The sidebar asks for a fresh sign-in link while it waits; the answer does
  // not change between two of those, so it is remembered for an hour.
  const cache = CacheService.getScriptCache();
  const remembered = (cache.get(RESOLVED_CLIENT_CACHE_KEY) || "").split("\n");
  if (remembered.length === 2 && remembered[0] === redirectUri && remembered[1]) return remembered[1];

  const clientId = lookUpOAuthClientId(redirectUri);
  cache.put(RESOLVED_CLIENT_CACHE_KEY, `${redirectUri}\n${clientId}`, SIGN_IN_TIMEOUT_SECONDS);
  return clientId;
}

function lookUpOAuthClientId(redirectUri: string): string {
  if (isRedirectRegistered(OAUTH_CLIENT_ID, redirectUri)) return OAUTH_CLIENT_ID;

  const properties = PropertiesService.getScriptProperties();
  try {
    const saved = JSON.parse(properties.getProperty(SELF_REGISTERED_CLIENT_PROPERTY) || "null") as {
      clientId?: unknown;
      redirectUri?: unknown;
    } | null;
    if (saved && typeof saved.clientId === "string" && saved.clientId && saved.redirectUri === redirectUri) {
      return saved.clientId;
    }
  } catch {
    // Unreadable: register again below.
  }

  const clientId = registerOAuthClient(redirectUri);
  properties.setProperty(SELF_REGISTERED_CLIENT_PROPERTY, JSON.stringify({ clientId, redirectUri }));
  return clientId;
}

/** True when Jobo has `redirectUri` registered for `clientId`. An unknown client answers 404. */
function isRedirectRegistered(clientId: string, redirectUri: string): boolean {
  let response: GoogleAppsScript.URL_Fetch.HTTPResponse;
  try {
    response = UrlFetchApp.fetch(OAUTH_CLIENT_LOOKUP_URL + buildQuery({ client_id: clientId, redirect_uri: redirectUri }), {
      method: "get",
      headers: { Accept: "application/json", "X-Jobo-Client": "sheets" },
      muteHttpExceptions: true,
    });
  } catch {
    throw new Error(SIGN_IN_UNAVAILABLE);
  }
  const status = response.getResponseCode();
  if (status === 404) return false;
  if (status < 200 || status >= 300) throw new Error(SIGN_IN_UNAVAILABLE);
  try {
    const body = JSON.parse(response.getContentText()) as { redirect_host?: unknown } | null;
    // The host is echoed back only for an address the client registered.
    return Boolean(body && typeof body.redirect_host === "string" && body.redirect_host);
  } catch {
    return false;
  }
}

function registerOAuthClient(redirectUri: string): string {
  let response: GoogleAppsScript.URL_Fetch.HTTPResponse;
  try {
    response = UrlFetchApp.fetch(OAUTH_REGISTER_URL, {
      method: "post",
      contentType: "application/json",
      headers: { Accept: "application/json", "X-Jobo-Client": "sheets" },
      payload: JSON.stringify({ client_name: SELF_REGISTERED_CLIENT_NAME, redirect_uris: [redirectUri] }),
      muteHttpExceptions: true,
    });
  } catch {
    throw new Error(SIGN_IN_UNAVAILABLE);
  }
  let clientId: unknown;
  try {
    clientId = (JSON.parse(response.getContentText()) as { client_id?: unknown } | null)?.client_id;
  } catch {
    clientId = undefined;
  }
  const status = response.getResponseCode();
  if (status < 200 || status >= 300 || typeof clientId !== "string" || !clientId) {
    throw new Error(SIGN_IN_UNAVAILABLE);
  }
  return clientId;
}

/**
 * Starts one sign-in attempt and returns the authorization URL for the sidebar
 * to show as a link.
 *
 * The verifier never leaves the server: it waits in the user cache under the
 * attempt's id, and that id travels inside the state token, which only Apps
 * Script can read or mint. The cache is per user and expires on its own, so an
 * abandoned attempt leaves nothing behind. (CacheService may evict early when a
 * lot is cached; that attempt then fails cleanly and the user signs in again.)
 */
export function startSignIn(): { url: string } {
  const clientId = resolveOAuthClientId(oauthRedirectUri());
  const attempt = Utilities.getUuid();
  const verifier = createCodeVerifier();
  CacheService.getUserCache().put(VERIFIER_CACHE_PREFIX + attempt, verifier, SIGN_IN_TIMEOUT_SECONDS);

  // The client rides in the state token too: the token request must name the
  // client the code was issued to, and only Apps Script can mint or alter a token.
  const state = ScriptApp.newStateToken()
    .withMethod(OAUTH_CALLBACK_FUNCTION)
    .withArgument("attempt", attempt)
    .withArgument("client", clientId)
    .withTimeout(SIGN_IN_TIMEOUT_SECONDS)
    .createToken();

  return {
    url:
      OAUTH_AUTHORIZE_URL +
      buildQuery({
        response_type: "code",
        client_id: clientId,
        redirect_uri: oauthRedirectUri(),
        scope: OAUTH_SCOPE,
        state,
        code_challenge: codeChallenge(verifier),
        code_challenge_method: "S256",
      }),
  };
}

/** What Apps Script hands the callback: the query parameters plus the state token's arguments. */
export interface SignInCallbackRequest {
  parameter?: Record<string, string | undefined>;
}

/** The tab the user lands on after Jobo redirects back. It never shows the key. */
function signInPage(ok: boolean, heading: string, message: string, hint: string): GoogleAppsScript.HTML.HtmlOutput {
  // Callback.html prints these with printing scriptlets, which escape them: the
  // message can be an error_description taken straight from the query string.
  const template = HtmlService.createTemplateFromFile("Callback");
  template.ok = ok;
  template.heading = heading;
  template.message = message;
  template.hint = hint;
  return template.evaluate().setTitle(CALLBACK_PAGE_TITLE);
}

function signInFailed(message: string, hint: string = TRY_AGAIN_HINT): GoogleAppsScript.HTML.HtmlOutput {
  return signInPage(false, "Sign-in failed", message, hint);
}

/**
 * Runs when Jobo redirects the browser back to /usercallback. Apps Script has
 * already checked the state token by the time this is called, and passes the
 * token's `attempt` argument alongside the query parameters.
 *
 * The attempt must still have a verifier waiting in this user's cache. A
 * callback for an unknown, expired or already finished attempt fails here, and
 * PKCE means a code issued for any other attempt would be refused by the token
 * endpoint anyway.
 */
export function authCallback(request: SignInCallbackRequest): GoogleAppsScript.HTML.HtmlOutput {
  const params = request?.parameter ?? {};

  // One use: whatever happens below, this attempt is over.
  const cache = CacheService.getUserCache();
  const attempt = params.attempt ?? "";
  const verifier = attempt ? cache.get(VERIFIER_CACHE_PREFIX + attempt) : null;
  if (attempt) cache.remove(VERIFIER_CACHE_PREFIX + attempt);

  if (params.error) {
    // access_denied is the user declining on the consent screen.
    const cancelled = params.error === "access_denied";
    const description =
      params.error_description ||
      (cancelled ? "Access to your Jobo account was not approved." : `Jobo returned the error ${params.error}.`);
    return signInPage(false, cancelled ? "Sign-in was cancelled" : "Sign-in failed", description, TRY_AGAIN_HINT);
  }
  if (!params.code) return signInFailed("Jobo did not return an authorization code.");
  if (!verifier) {
    return signInFailed(
      "This sign-in link has expired or was already used.",
      "Close this tab and check the sidebar. If it is not connected, choose Sign in with Jobo again.",
    );
  }

  // UrlFetchApp throws on a network fault even with muteHttpExceptions.
  let response: GoogleAppsScript.URL_Fetch.HTTPResponse;
  try {
    response = UrlFetchApp.fetch(OAUTH_TOKEN_URL, {
      method: "post",
      contentType: "application/x-www-form-urlencoded",
      headers: { Accept: "application/json", "X-Jobo-Client": "sheets" },
      // A public client: sending a client_secret makes the server answer invalid_client.
      payload: buildQuery({
        grant_type: "authorization_code",
        client_id: params.client || OAUTH_CLIENT_ID,
        redirect_uri: oauthRedirectUri(),
        code: params.code,
        code_verifier: verifier,
      }).slice(1),
      muteHttpExceptions: true,
    });
  } catch {
    return signInFailed("Jobo could not be reached to finish signing in.");
  }

  let body: Record<string, unknown> = {};
  try {
    const parsed: unknown = JSON.parse(response.getContentText());
    if (parsed && typeof parsed === "object") body = parsed as Record<string, unknown>;
  } catch {
    // Not JSON: handled below as a failed exchange or a missing key.
  }

  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    const description = typeof body.error_description === "string" ? body.error_description : "";
    return signInFailed(description || `Jobo returned HTTP ${status}.`);
  }

  // The access token is a Jobo API key, so it is stored exactly like a pasted one.
  const saved = saveApiKey(typeof body.access_token === "string" ? body.access_token : "");
  if (!saved.ok) {
    return signInFailed(
      "Jobo did not return a valid API key.",
      "Close this tab and try again from the sidebar. You can also paste an API key there.",
    );
  }

  const userProperties = PropertiesService.getUserProperties();
  // The app the key was issued to; see disconnect().
  userProperties.setProperty(OAUTH_CLIENT_PROPERTY, params.client || OAUTH_CLIENT_ID);

  // Present only when the email scope was granted.
  const email = typeof body.account_email === "string" ? body.account_email.trim() : "";
  if (email) userProperties.setProperty(ACCOUNT_EMAIL_PROPERTY, email);

  return signInPage(
    true,
    "Connected to Jobo",
    email ? `Your Jobo account ${email} is connected.` : "Your Jobo account is connected.",
    "You can close this tab and return to your spreadsheet.",
  );
}

// v2: added skills and the job category/subcategory options.
const FILTER_OPTIONS_CACHE_KEY = "jobo_filter_options_v2";
/** 6 hours — the CacheService maximum. Source/industry churn is slower than this. */
const FILTER_OPTIONS_CACHE_SECONDS = 21600;

/** A picker option: `value` is what the filter takes, `label` what the user reads. */
export interface PickerOption {
  value: string;
  label: string;
  count?: number;
}

export interface FilterOptionsResult {
  ok: boolean;
  sources: string[];
  industries: string[];
  /** The most common skills; the skills pickers still accept any typed skill. */
  skills: string[];
  jobCategories: PickerOption[];
  jobSubcategories: PickerOption[];
}

const NO_FILTER_OPTIONS: FilterOptionsResult = {
  ok: false,
  sources: [],
  industries: [],
  skills: [],
  jobCategories: [],
  jobSubcategories: [],
};

function readOptions(body: unknown, listField: string): JobFacet[] {
  if (!body || typeof body !== "object") return [];
  const list = (body as Record<string, unknown>)[listField];
  if (!Array.isArray(list)) return [];
  return list.filter(
    (entry): entry is JobFacet =>
      Boolean(entry) && typeof (entry as JobFacet).key === "string" && (entry as JobFacet).key.length > 0,
  );
}

function readOptionKeys(body: unknown, listField: string): string[] {
  return readOptions(body, listField).map((option) => option.key);
}

function titleCase(value: string): string {
  return value.replace(/(^|[\s/&-])([a-z])/g, (_match, lead: string, letter: string) => lead + letter.toUpperCase());
}

/**
 * Category / subcategory picker options from the occupation facets. The value
 * is the display name ("Healthcare"), which the case-insensitive filter takes;
 * a subcategory's label carries its category, since names alone are terse.
 */
export function occupationLevelOptions(options: JobFacet[], withParent: boolean): PickerOption[] {
  return options
    .map((option) => {
      const name = option.name ?? titleCase(option.key);
      const label = withParent && option.parent ? `${name} · ${titleCase(option.parent)}` : name;
      return { value: name, label, count: option.count };
    })
    .sort((a, b) => a.label.localeCompare(b.label));
}

/**
 * Dynamic filter options for the sidebar comboboxes. Every list is small
 * enough (≤250) to ship whole and filter client-side — a per-keystroke
 * google.script.run round trip is far too slow for typeahead. Cached in the
 * script cache (shared across users; the lists are not user-specific).
 * `ok: false` tells the sidebar to fall back to free-text chips. The
 * occupation facets answer 404 until occupations are published; those lists
 * then stay empty and their pickers take free text.
 */
function getFilterOptions(): FilterOptionsResult {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(FILTER_OPTIONS_CACHE_KEY);
    if (cached) return JSON.parse(cached) as FilterOptionsResult;

    const key = apiKey();
    const filters = request("/api/connectors/filters", key);
    if (filters.status < 200 || filters.status >= 300) return NO_FILTER_OPTIONS;

    const facet = (name: string): JobFacet[] => {
      const result = request(`/api/connectors/filter-options/${name}?limit=250`, key);
      return result.status >= 200 && result.status < 300 ? readOptions(result.body, "options") : [];
    };

    const result: FilterOptionsResult = {
      ok: true,
      sources: readOptionKeys(filters.body, "sources"),
      industries: facet("industries").map((option) => option.key),
      skills: facet("skills").map((option) => option.key),
      jobCategories: occupationLevelOptions(facet("job_categories"), false),
      jobSubcategories: occupationLevelOptions(facet("job_subcategories"), true),
    };
    try {
      cache.put(FILTER_OPTIONS_CACHE_KEY, JSON.stringify(result), FILTER_OPTIONS_CACHE_SECONDS);
    } catch {
      // Over CacheService's 100 KB value limit: serve it uncached.
    }
    return result;
  } catch {
    return NO_FILTER_OPTIONS;
  }
}

/**
 * Every occupation, flattened for the occupation pickers: `value` is the id
 * `occupation_ids` takes, `label` the name with its category › subcategory.
 * The sidebar loads it once and filters locally (a few thousand entries —
 * too large for CacheService, so fetched per sidebar open; the endpoint is
 * free). Empty when occupations are not published or the call fails, which
 * leaves the pickers on free text.
 */
function getOccupations(): PickerOption[] {
  try {
    const result = request("/api/occupations", apiKey());
    if (result.status < 200 || result.status >= 300) return [];
    return flattenOccupations(result.body as OccupationTaxonomyResponse);
  } catch {
    return [];
  }
}

export function flattenOccupations(taxonomy: OccupationTaxonomyResponse | null | undefined): PickerOption[] {
  const out: PickerOption[] = [];
  for (const category of taxonomy?.categories ?? []) {
    for (const sub of category.subcategories ?? []) {
      for (const occupation of sub.occupations ?? []) {
        out.push({ value: occupation.id, label: `${occupation.name} (${category.name} › ${sub.name})` });
      }
    }
  }
  return out.sort((a, b) => a.label.localeCompare(b.label));
}

export interface CompanySuggestionResult {
  suggestions: Array<{ label: string; value: string; ids: string[]; count?: number }>;
}

/**
 * Company typeahead proxy (name or domain prefix, 2+ characters). Each pick
 * carries every board id of the company, which is what the `companies`
 * filter should receive — a name would also match every other company called
 * that. Never throws: an empty list keeps the field usable as free text.
 */
function suggestCompanies(q: string): CompanySuggestionResult {
  const trimmed = (q || "").trim();
  if (trimmed.length < 2) return { suggestions: [] };
  try {
    const result = request(
      "/api/connectors/companies/suggest" + buildQuery({ q: trimmed, limit: 8 }),
      apiKey(),
    );
    if (result.status < 200 || result.status >= 300) return { suggestions: [] };
    const body = result.body as { suggestions?: CompanySuggestion[] };
    return { suggestions: (Array.isArray(body.suggestions) ? body.suggestions : []).map(companyPickerItem) };
  } catch {
    return { suggestions: [] };
  }
}

export function companyPickerItem(company: CompanySuggestion): CompanySuggestionResult["suggestions"][number] {
  const ids = company.ids && company.ids.length > 0 ? company.ids : [company.id];
  const details = [company.domain, company.country].filter(Boolean).join(", ");
  return {
    label: details ? `${company.name} (${details})` : company.name,
    // Unique per company, so picking it twice is a no-op in the chip row.
    value: ids.join(","),
    ids,
    ...(typeof company.job_count === "number" ? { count: company.job_count } : {}),
  };
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
 * add noise without a use case (posted_after alone covers "recent jobs") —
 * and description_format, since the sheet has no description column.
 */
export const SHEETS_SUPPORTED_FILTER_KEYS: readonly string[] = [
  "q",
  "location",
  "sources",
  "companies",
  "exclude_companies",
  "work_model",
  "employment_type",
  "experience_level",
  "language",
  "skills",
  "exclude_skills",
  "industries",
  "exclude_industries",
  "company_categories",
  "exclude_company_categories",
  "job_categories",
  "exclude_job_categories",
  "job_subcategories",
  "exclude_job_subcategories",
  "occupation_ids",
  "exclude_occupation_ids",
  "min_salary_usd",
  "max_salary_usd",
  "posted_after",
  "search_description",
  "sort",
];

/**
 * What the sidebar sends. Every field is optional and an absent or empty one
 * is no filter, so a request from an older sidebar (or any subset of fields)
 * keeps working.
 */
export interface ImportRequest {
  q?: string;
  location?: string;
  sources?: string[];
  /** Company ids (every id of a picked company), domains or whole names. */
  companies?: string[];
  excludeCompanies?: string[];
  skills?: string[];
  excludeSkills?: string[];
  industries?: string[];
  excludeIndustries?: string[];
  companyCategories?: string[];
  excludeCompanyCategories?: string[];
  jobCategories?: string[];
  excludeJobCategories?: string[];
  jobSubcategories?: string[];
  excludeJobSubcategories?: string[];
  occupationIds?: string[];
  excludeOccupationIds?: string[];
  languages?: string[];
  workModels?: string[];
  employmentTypes?: string[];
  experienceLevels?: string[];
  minSalaryUsd?: number;
  maxSalaryUsd?: number;
  postedAfter?: string;
  searchDescriptions?: boolean;
  sort?: string;
  limit?: number;
}

const SORTS: readonly SearchSortValue[] = ["relevance", "discovered_asc", "discovered_desc"];

/** Sidebar request → connector-core filter values (keyed by JOB_SEARCH_FILTERS keys). */
export function toSearchFilters(req: ImportRequest): JobSearchFilters {
  const sort = SORTS.find((value) => value === req.sort);
  return {
    q: req.q,
    location: req.location,
    sources: req.sources,
    companies: req.companies,
    exclude_companies: req.excludeCompanies,
    work_model: req.workModels,
    employment_type: req.employmentTypes,
    experience_level: req.experienceLevels,
    language: req.languages,
    skills: req.skills,
    exclude_skills: req.excludeSkills,
    industries: req.industries,
    exclude_industries: req.excludeIndustries,
    company_categories: req.companyCategories,
    exclude_company_categories: req.excludeCompanyCategories,
    job_categories: req.jobCategories,
    exclude_job_categories: req.excludeJobCategories,
    job_subcategories: req.jobSubcategories,
    exclude_job_subcategories: req.excludeJobSubcategories,
    occupation_ids: req.occupationIds,
    exclude_occupation_ids: req.excludeOccupationIds,
    min_salary_usd: req.minSalaryUsd,
    max_salary_usd: req.maxSalaryUsd,
    posted_after: req.postedAfter || undefined,
    // true is the API default; only the opt-out travels on the wire.
    search_description: req.searchDescriptions === false ? false : undefined,
    sort,
  };
}

/** The API's filter warnings (e.g. a company name that matched nothing), as one sentence each. */
function warningText(body: unknown): string[] {
  const warnings = (body as JobSearchResponse | null)?.warnings;
  return Array.isArray(warnings)
    ? warnings.map((warning) => warning?.message).filter((m): m is string => typeof m === "string" && m.length > 0)
    : [];
}

/**
 * One page size for the whole import (the API caps it at 100): the fewest
 * pages, split evenly, so the last page overshoots `limit` by at most one job
 * per page — 150 → 2 × 75, 1000 → 10 × 100, 101 → 2 × 51.
 */
export function importPageSize(limit: number): number {
  return Math.ceil(limit / Math.ceil(limit / 100));
}

/**
 * How long an import keeps starting new work: about 4.5 of Apps Script's 6
 * minutes per run. A page started just before it, with a retry, still ends
 * well inside the limit; a run that hits the limit is killed mid-page.
 */
export const IMPORT_TIME_BUDGET_MS = 270_000;

/**
 * Every page is billed as it is delivered, so each page is written to the sheet
 * as soon as it arrives. A failure on a later page (a 402 once the wallet runs
 * dry, a 429 that outlasted the retries, or a search that cannot be retried)
 * keeps the pages already written, and the message says how far the import got
 * and why it stopped. So does running out of time: the import stops starting
 * pages after IMPORT_TIME_BUDGET_MS rather than be killed with a page billed
 * and not written.
 */
export function importJobs(req: ImportRequest): { ok: boolean; message: string } {
  try {
    const key = apiKey();
    const limit = Math.min(Math.max(req.limit ?? 100, 1), 1000);
    const deadline = Date.now() + IMPORT_TIME_BUDGET_MS;

    const sheet = sheetWriter(COLUMNS, "Jobo Jobs");
    let page = 1;
    const usage = newUsage();
    let failure: string | null = null;
    let outOfTime = false;
    const notes: string[] = [];
    const filters = toSearchFilters(req);
    const pageSize = importPageSize(limit);

    while (sheet.written() < limit) {
      if (page > 1 && Date.now() >= deadline) {
        outOfTime = true;
        break;
      }

      // POST /api/jobs/search: the include/exclude and taxonomy filters exist
      // only on the body surface. toJobSearchBody drops empty values, so an
      // untouched picker never becomes a filter.
      const body = toJobSearchBody(filters, {
        page,
        // Constant across pages: `page` offsets by page_size, so shrinking
        // the last page re-fetched jobs already imported.
        page_size: pageSize,
        // The sheet has no description column; facets are never shown.
        include_facets: [],
        include_fields: [],
      });

      // The search is billed per job returned, so it is never replayed after a
      // lost answer (see request()). A page that got no answer at all throws,
      // and stops the loop the same way a failed status does.
      let result: ApiResult;
      try {
        result = request("/api/jobs/search", key, body, { replaySafe: false, deadline });
      } catch (err) {
        failure = (err as Error).message;
        break;
      }
      if (result.status < 200 || result.status >= 300) {
        failure = describeFailure(result.status, result.body);
        break;
      }

      addUsage(usage, result);

      if (page === 1) notes.push(...warningText(result.body));

      const data = result.body as { jobs?: Job[]; total_pages?: number };
      const batch = data.jobs ?? [];
      // Equal pages can overshoot the limit by a few jobs; never write more.
      try {
        sheet.append(batch.slice(0, limit - sheet.written()));
      } catch (err) {
        failure = `Writing to the sheet failed: ${(err as Error).message}`;
        break;
      }

      if (batch.length === 0 || page >= (data.total_pages ?? 1)) break;
      page++;
    }

    const written = sheet.written();
    const note = notes.length > 0 ? ` Note: ${notes.join(" ")}` : "";
    if (written === 0) {
      return failure != null
        ? { ok: false, message: failure }
        : { ok: true, message: `No jobs matched those filters.${note}` };
    }

    sheet.finish();

    const wallet = walletText(usage);
    const included = usage.quotaRemaining != null
      ? ` Included jobs: ${usage.quotaRemaining}${usage.quotaLimit != null ? ` of ${usage.quotaLimit}` : ""} remaining.`
      : "";
    if (failure != null) {
      return {
        ok: false,
        message: `Imported ${written} of the ${limit} jobs requested before the import stopped. ${failure}${wallet}${included}${note}`,
      };
    }
    if (outOfTime) {
      return {
        ok: false,
        message:
          `Imported ${written} of the ${limit} jobs requested. The import stopped early to finish inside ` +
          `Google's 6-minute limit for one run. Importing again starts from the first matching job and bills ` +
          `again for the jobs already imported.${wallet}${included}${note}`,
      };
    }
    return { ok: true, message: `Imported ${written} jobs.${wallet}${included}${note}` };
  } catch (err) {
    return { ok: false, message: (err as Error).message };
  }
}

type Sheet = GoogleAppsScript.Spreadsheet.Sheet;

/**
 * The sheet one import writes to. Without `existing`, a new tab named after
 * `baseName` is created with its header row when the first item arrives; with
 * it, rows are appended below what the tab already holds. Each append is one
 * setValues call followed by a flush, so a page is in the spreadsheet before
 * the next one is requested: a run that is killed or fails later keeps every
 * page it was billed for.
 */
function sheetWriter<T>(
  columns: Array<Column<T>>,
  baseName: string,
  existing: Sheet | null = null,
): { append: (items: T[]) => void; finish: () => void; written: () => number; name: () => string } {
  let sheet: Sheet | null = existing;
  const startRow = existing ? existing.getLastRow() + 1 : 1;
  let rowsWritten = 0;
  let written = 0;

  return {
    written: () => written,
    name: () => (sheet ? sheet.getName() : ""),
    append(items: T[]): void {
      if (items.length === 0) return;
      const rows: Array<Array<string | number>> = [];
      // Header and first page in one write; formatted now, not at the end, so
      // a sheet left by a killed run still reads as a table.
      if (!sheet) rows.push(columns.map((c) => c.header));
      for (const item of items) rows.push(columns.map((c) => neutralizeFormula(c.value(item))));

      if (!sheet) {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        sheet = spreadsheet.insertSheet(uniqueSheetName(spreadsheet, baseName));
        sheet.setFrozenRows(1);
      }
      const firstRow = startRow + rowsWritten;
      // Before the values: a value is read as a date or a number as it lands.
      columns.forEach((column, index) => {
        if (column.text) (sheet as Sheet).getRange(firstRow, index + 1, rows.length, 1).setNumberFormat("@");
      });
      sheet.getRange(firstRow, 1, rows.length, columns.length).setValues(rows);
      if (firstRow === 1) sheet.getRange(1, 1, 1, columns.length).setFontWeight("bold");
      SpreadsheetApp.flush();
      rowsWritten += rows.length;
      written += items.length;
    },
    finish(): void {
      // Once, at the end: resizing measures every row written.
      if (sheet) sheet.autoResizeColumns(1, columns.length);
    },
  };
}

/** What the billing headers said over one import: the charges summed, the latest balance and allowance. */
interface Usage {
  creditsSpent: number;
  balance: number | null;
  quotaLimit: number | null;
  quotaRemaining: number | null;
  /** True once any response carried a billing header. The Unlimited plan sends none. */
  reported: boolean;
}

function newUsage(): Usage {
  return { creditsSpent: 0, balance: null, quotaLimit: null, quotaRemaining: null, reported: false };
}

function addUsage(usage: Usage, result: ApiResult): void {
  usage.creditsSpent += result.creditsDeducted ?? 0;
  usage.balance = result.creditsBalance ?? usage.balance;
  usage.quotaLimit = result.quotaLimit ?? usage.quotaLimit;
  usage.quotaRemaining = result.quotaRemaining ?? usage.quotaRemaining;
  if ([result.creditsDeducted, result.creditsBalance, result.quotaLimit, result.quotaRemaining].some((v) => v != null)) {
    usage.reported = true;
  }
}

/** " Cost: $0.30 (300 credits). Wallet balance: $4.20." — each part only when the headers said it. */
function walletText(usage: Usage): string {
  const cost =
    usage.creditsSpent > 0 ? ` Cost: $${(usage.creditsSpent / 1000).toFixed(2)} (${usage.creditsSpent} credits).` : "";
  const left = usage.balance != null ? ` Wallet balance: $${(usage.balance / 1000).toFixed(2)}.` : "";
  return cost + left;
}

// ── Hiring Signals ───────────────────────────────────────────────────────────
//
// GET /api/signals, read into a tab of its own. It is built from the job
// import's pieces: request() with its retry policy, marked as billed so a lost
// answer is never replayed; the same sheet writer, formula neutralising and
// time budget. What differs: paging is by cursor; a watchlist of more than
// 100 domains is read in requests of at most 100 that share one budget; and
// "only new" appends to a signals tab, reading from its newest signal on.
//
// connector-core's JoboClient and pollSignals are async and cannot run here
// (see the top of this file); its pure helpers are used as they are:
// toSignalsQuery for the query, describeSignal and the labels for the rows.

/** The most signals one import reads. Each one is billed, at ten times a job. */
export const SIGNALS_IMPORT_MAX = 1000;

/**
 * The signal types as the dashboard groups them. A type missing here (one the
 * core learns later) is still offered, under "Other": see signalTypeGroups().
 */
export const SIGNAL_TYPE_GROUPS: ReadonlyArray<{ label: string; types: readonly string[] }> = [
  { label: "Growth", types: ["hiring_surge", "function_surge", "hiring_resumed"] },
  { label: "Expansion", types: ["new_country", "new_city", "new_function", "first_role"] },
  { label: "Leadership", types: ["leadership_hire", "leadership_role_closed"] },
  { label: "Tech and skills", types: ["new_technology", "new_skill"] },
  { label: "Slowing or struggling", types: ["hiring_slowdown", "hiring_freeze", "hard_to_fill"] },
];

export interface SignalTypeGroup {
  label: string;
  types: Array<{ value: string; label: string }>;
}

/** The sidebar's signal type checkboxes, with the core's labels. */
export function signalTypeGroups(): SignalTypeGroup[] {
  const option = (type: string) => ({ value: type, label: signalTypeLabel(type) });
  const groups: SignalTypeGroup[] = SIGNAL_TYPE_GROUPS.map((group) => ({
    label: group.label,
    types: group.types.map(option),
  }));
  const grouped = new Set(SIGNAL_TYPE_GROUPS.flatMap((group) => group.types));
  const others = SIGNAL_TYPES.filter((type) => !grouped.has(type));
  if (others.length > 0) groups.push({ label: "Other", types: others.map(option) });
  return groups;
}

/** A list cell: the values joined with ", ", ignoring anything that is not text. */
function listCell(values: unknown): string {
  return Array.isArray(values) ? values.filter((v) => typeof v === "string" && v !== "").join(", ") : "";
}

/**
 * Columns of a signals tab, in order. "Only new" finds its way back by these
 * headers, so renaming one makes earlier tabs unrecognised.
 *
 * Every field is read defensively: a type added after this was written may
 * carry less, and describeSignal() falls back to the type's label.
 */
export const SIGNAL_COLUMNS: Array<Column<HiringSignal>> = [
  { header: "Detected", value: (s) => s.detected_on ?? "" },
  { header: "Company", value: (s) => s.company?.name ?? "" },
  { header: "Domain", value: (s) => s.company?.domain ?? "" },
  { header: "Signal", value: (s) => signalTypeLabel(s.type) },
  { header: "Summary", value: (s) => describeSignal(s) },
  { header: "Strength", value: (s) => (typeof s.strength === "number" ? s.strength : "") },
  { header: "Strength band", value: (s) => (typeof s.strength === "number" ? signalStrengthBand(s.strength) : "") },
  { header: "Country", value: (s) => s.company?.country_code ?? "" },
  // Text: Sheets reads a band like "1-10" as a date.
  { header: "Company size", value: (s) => s.company?.size ?? "", text: true },
  { header: "Industry", value: (s) => s.company?.industry ?? "" },
  { header: "Funding stage", value: (s) => s.company?.funding_stage ?? "" },
  { header: "Functions", value: (s) => listCell(s.focus?.functions) },
  { header: "Role", value: (s) => s.focus?.title || s.focus?.occupation || "" },
  { header: "Seniority", value: (s) => s.focus?.seniority ?? "" },
  { header: "Topics", value: (s) => listCell(s.focus?.topics) },
  { header: "Job countries", value: (s) => listCell(s.focus?.job_countries) },
  { header: "Evidence job IDs", value: (s) => listCell(s.evidence_job_ids) },
  // Text, so the seven decimal places "only new" sends back are kept as they came.
  { header: "Created at", value: (s) => s.created_at ?? "", text: true },
  { header: "Signal ID", value: (s) => s.id ?? "", text: true },
];

const CREATED_AT_COLUMN = SIGNAL_COLUMNS.findIndex((c) => c.header === "Created at") + 1;
const SIGNAL_ID_COLUMN = SIGNAL_COLUMNS.findIndex((c) => c.header === "Signal ID") + 1;

/**
 * One domain as the API stores it: the host, lower-cased, without a scheme,
 * `www.`, path, user or port. Mirrors HiringSignalDomains.Normalize on the
 * server, which refuses the whole request for a value it cannot read; null
 * here is a value that is not a domain.
 */
export function normalizeDomain(value: string): string | null {
  let host = String(value ?? "").trim().toLowerCase();
  const scheme = host.indexOf("://");
  if (scheme >= 0) host = host.slice(scheme + 3);
  const end = host.search(/[/?#]/);
  if (end >= 0) host = host.slice(0, end);
  host = host.slice(host.lastIndexOf("@") + 1);
  const port = host.indexOf(":");
  if (port >= 0) host = host.slice(0, port);
  if (host.startsWith("www.")) host = host.slice(4);
  host = host.replace(/^\.+|\.+$/g, "");
  return host.length === 0 || !host.includes(".") || /\s/.test(host) ? null : host;
}

/**
 * Website domains from the sidebar's box or from cells: split on line breaks,
 * commas, semicolons and spaces, normalised, de-duplicated in the order given.
 * `rejected` holds what is not a domain, such as a header cell.
 */
export function cleanDomains(input: string | ReadonlyArray<unknown> | null | undefined): {
  domains: string[];
  rejected: string[];
} {
  const values: ReadonlyArray<unknown> = typeof input === "string" ? [input] : input ?? [];
  const domains: string[] = [];
  const rejected: string[] = [];
  for (const value of values) {
    if (value == null) continue;
    for (const part of String(value).split(/[\s,;]+/)) {
      if (part === "") continue;
      const domain = normalizeDomain(part);
      if (domain == null) {
        if (!rejected.includes(part)) rejected.push(part);
      } else if (!domains.includes(domain)) {
        domains.push(domain);
      }
    }
  }
  return { domains, rejected };
}

/** The sidebar's "Use the selected cells": the domains in the selection, and how many cells held none. */
export function getSelectedDomains(): { domains: string[]; skipped: number } {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) return { domains: [], skipped: 0 };
  // Display values: a cell showing a link gives the text it shows.
  const { domains, rejected } = cleanDomains(range.getDisplayValues().flat());
  return { domains, skipped: rejected.length };
}

/** An instant as `created_at` writes it, split so that all seven decimal places compare. */
interface Instant {
  /** Whole seconds since the epoch. */
  seconds: number;
  /** The decimal places, padded to seven. */
  fraction: string;
}

/**
 * Reads `2026-11-20T09:15:30.1234567Z` (a fraction of up to seven places, and
 * Z, an offset or neither, which the API reads as UTC). A JavaScript Date
 * keeps milliseconds only, so the fraction is kept as text.
 */
export function parseInstant(text: unknown): Instant | null {
  const match = /^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})(?:\.(\d{1,7}))?(Z|[+-]\d{2}:\d{2})?$/.exec(String(text ?? "").trim());
  if (!match) return null;
  const ms = Date.parse(`${match[1]}${match[3] ?? "Z"}`);
  if (!isFinite(ms)) return null;
  return { seconds: Math.floor(ms / 1000), fraction: (match[2] ?? "").padEnd(7, "0") };
}

function compareInstants(a: Instant, b: Instant): number {
  if (a.seconds !== b.seconds) return a.seconds - b.seconds;
  return a.fraction < b.fraction ? -1 : a.fraction > b.fraction ? 1 : 0;
}

/** The newest of these `created_at` values, as written; null when none can be read. */
export function newestCreatedAt(values: ReadonlyArray<unknown>): string | null {
  let newest: { text: string; instant: Instant } | null = null;
  for (const value of values) {
    const instant = parseInstant(value);
    if (instant && (!newest || compareInstants(instant, newest.instant) > 0)) {
      newest = { text: String(value).trim(), instant };
    }
  }
  return newest ? newest.text : null;
}

/**
 * Where an "only new" import starts reading: the newest Created at in the
 * tab, exactly as it is written there.
 *
 * "Only new" reads in the order signals were stored (`sort=created_asc`).
 * There `created_after` is exclusive and compares every decimal place, and a
 * signal that becomes readable later never has an earlier `created_at`. So
 * the newest value, sent back as the string the tab holds, reads nothing the
 * tab already has and misses nothing stored since: no overlap, and nothing
 * billed twice. The value is never parsed into a date on the way: a date
 * keeps milliseconds, and the signals of one batch are a microsecond apart.
 *
 * Null: the tab holds no signal to read after.
 */
export function signalsReadBound(newest: string | null): string | null {
  if (newest == null || parseInstant(newest) == null) return null;
  return newest.trim();
}

/** A tab this add-on wrote signals to, with what "only new" needs from it. */
export interface SignalsTab {
  sheet: Sheet;
  name: string;
  /** Signal rows below the header. */
  rows: number;
  ids: Set<string>;
  /** Null when the tab has no Created at value that can be read. */
  newestCreatedAt: string | null;
}

/** A tab is a signals tab when its header row starts with SIGNAL_COLUMNS' headers, in order. */
export function readSignalsTab(sheet: Sheet | null): SignalsTab | null {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < SIGNAL_COLUMNS.length) return null;
  const header = sheet.getRange(1, 1, 1, SIGNAL_COLUMNS.length).getDisplayValues()[0];
  if (!SIGNAL_COLUMNS.every((column, index) => String(header[index]).trim() === column.header)) return null;

  const rows = sheet.getLastRow() - 1;
  const column = (index: number): string[] =>
    rows > 0 ? sheet.getRange(2, index, rows, 1).getDisplayValues().map((row) => String(row[0]).trim()) : [];
  return {
    sheet,
    name: sheet.getName(),
    rows,
    ids: new Set(column(SIGNAL_ID_COLUMN).filter((id) => id !== "")),
    newestCreatedAt: newestCreatedAt(column(CREATED_AT_COLUMN)),
  };
}

/** For the sidebar: whether the open tab is a signals tab, so "only new" can be offered. */
export function getSignalsTabStatus(): { isSignalsTab: boolean; name: string; rows: number } {
  try {
    const tab = readSignalsTab(SpreadsheetApp.getActiveSheet());
    return tab ? { isSignalsTab: true, name: tab.name, rows: tab.rows } : { isSignalsTab: false, name: "", rows: 0 };
  } catch {
    return { isSignalsTab: false, name: "", rows: 0 };
  }
}

/**
 * What the signals view sends. Every field is optional and an empty one is
 * no filter.
 */
export interface SignalsImportRequest {
  /** Website domains: the sidebar's box as typed, or a list. */
  domains?: string | string[];
  types?: string[];
  /** Two-letter codes of where the companies are headquartered. */
  countries?: string[];
  companySizes?: string[];
  /** Job categories the signals must be about, e.g. Technology. */
  functions?: string[];
  minStrength?: number;
  /** yyyy-MM-dd, inclusive. */
  detectedAfter?: string;
  excludeAgencies?: boolean;
  /** Append to the open signals tab, reading only what is new since its newest signal. */
  onlyNew?: boolean;
  limit?: number;
}

/**
 * HIRING_SIGNAL_FILTERS keys this connector sends, checked against the core by
 * a test. The rest stay in the dashboard and the API: a sheet is for a list
 * of companies and what changed in their hiring, not for every focus filter.
 */
export const SHEETS_SIGNAL_FILTER_KEYS: readonly string[] = [
  "types",
  "domains",
  "countries",
  "company_sizes",
  "functions",
  "min_strength",
  "detected_after",
  "exclude_agencies",
  "created_after",
];

/** Sidebar request → connector-core signal filters (without the domains, which are read in chunks). */
export function toSignalFilters(req: SignalsImportRequest): HiringSignalFilters {
  return {
    types: req.types,
    countries: (req.countries ?? []).map((code) => String(code ?? "").trim().toUpperCase()),
    company_sizes: req.companySizes,
    functions: req.functions,
    min_strength: req.minStrength,
    detected_after: req.detectedAfter || undefined,
    exclude_agencies: req.excludeAgencies === true,
  };
}

/** True when a filter other than the domains says which companies or subjects are wanted. */
function hasNarrowingFilter(filters: HiringSignalFilters): boolean {
  const query = toSignalsQuery(filters);
  return NARROWING_SIGNAL_FILTER_KEYS.some((key) => query[key] !== undefined);
}

/** Human-readable message for a failed signals read. */
export function describeSignalsFailure(status: number, body: unknown): string {
  const fields = body && typeof body === "object" ? (body as Record<string, unknown>) : {};
  const code = typeof fields.code === "string" ? fields.code : null;

  if (status === 402) {
    return (
      "Your plan's included signals and your Jobo wallet do not cover the next page of signals, so it was not " +
      "read and nothing was charged for it. Top up your wallet at https://enterprise.jobo.world/credits?src=sheets, " +
      "or import fewer signals."
    );
  }
  if (status === 403 && code === "account_key_required") {
    return (
      "Hiring Signals is billed to a Jobo account, and this API key does not belong to one. Click Disconnect at " +
      "the bottom of the sidebar, then sign in with Jobo or paste a key from https://enterprise.jobo.world/api-keys?src=sheets."
    );
  }
  if (status === 404 && code === "not_available_in_sandbox") {
    return (
      "Hiring Signals is not available with a sandbox key. Click Disconnect at the bottom of the sidebar, then " +
      "sign in with Jobo or paste a live key (jbe_live_…)."
    );
  }
  if (status === 429) {
    return (
      "Jobo rate limit reached. Without a subscription, Hiring Signals takes 20 requests a minute, 200 an hour " +
      "and 1,000 a day. Wait a minute and try again."
    );
  }
  if (status === 400 && fields.errors && typeof fields.errors === "object") {
    // A validation problem: the first message names the filter and what it takes.
    for (const messages of Object.values(fields.errors as Record<string, unknown>)) {
      if (Array.isArray(messages) && typeof messages[0] === "string") return messages[0];
    }
  }
  return describeFailure(status, body);
}

function counted(count: number, noun: string): string {
  return `${count} ${noun}${count === 1 ? "" : "s"}`;
}

/** What was billed, and what is left of the included signals and the wallet. */
function signalsUsageText(read: number, usage: Usage): string {
  if (read === 0) return " Nothing was billed.";
  const billed = usage.reported
    ? ` Billed for ${counted(read, "signal")}.`
    : ` Read ${counted(read, "signal")}; Jobo reported no charge for them.`;
  const included =
    usage.quotaRemaining != null
      ? ` Included signals left this month: ${usage.quotaRemaining}${usage.quotaLimit != null ? ` of ${usage.quotaLimit}` : ""}.`
      : "";
  return `${billed}${walletText(usage)}${included}`;
}

/**
 * Reads hiring signals into a new "Jobo Signals" tab, or with `onlyNew` adds
 * the new ones to the open signals tab.
 *
 * `limit` caps the signals *read*, which is what is billed. Each request asks
 * for no more than is left of it, so an import never reads past the cap.
 * Domains are read in chunks of at most 100 (the API's limit per list), one
 * after another, sharing that cap.
 *
 * An import into a new tab reads the newest detection day first. "Only new"
 * reads the stored order, oldest first, from the tab's newest signal on: with
 * one chunk of domains, an import that stops early is carried on exactly by
 * the next one, because what it wrote is everything stored up to its last row.
 *
 * As for jobs, each page is written as it arrives, a lost answer is not asked
 * for again, and no request is started after IMPORT_TIME_BUDGET_MS.
 */
export function importSignals(req: SignalsImportRequest): { ok: boolean; message: string } {
  try {
    const key = apiKey();
    const limit = Math.min(Math.max(Math.trunc(req.limit ?? 0) || SIGNAL_DEFAULT_LIMIT, 1), SIGNALS_IMPORT_MAX);
    const deadline = Date.now() + IMPORT_TIME_BUDGET_MS;

    const { domains, rejected } = cleanDomains(req.domains);
    if (rejected.length > 0) {
      return {
        ok: false,
        message: `Not a website domain: ${rejected.slice(0, 5).join(", ")}. Enter domains such as stripe.com, one per line or separated by commas.`,
      };
    }
    const filters = toSignalFilters(req);
    const badCountry = (filters.countries ?? []).find((code) => code !== "" && !/^[A-Z]{2}$/.test(code));
    if (badCountry !== undefined) {
      return { ok: false, message: `Company countries take two-letter codes, such as US or DE. "${badCountry}" is not one.` };
    }
    // Builds the query once, so a value that contains a comma is refused
    // (by toSignalsQuery, with the filter's name) before anything is read.
    const narrowed = hasNarrowingFilter(filters);
    if (domains.length === 0 && !narrowed) {
      return {
        ok: false,
        message:
          "Add the website domains of the companies you want, or choose at least one signal type, company " +
          "country, company size or function. Without one the import reads signals from every company Jobo " +
          "tracks, and each signal is billed.",
      };
    }

    let tab: SignalsTab | null = null;
    // True for "only new" with a signal in the tab to read after. A signals
    // tab with no rows yet is read like a new one, newest detection day
    // first: the stored order without a bound starts at the oldest signal
    // Jobo holds.
    let storedOrder = false;
    if (req.onlyNew) {
      tab = readSignalsTab(SpreadsheetApp.getActiveSheet());
      if (!tab) {
        return {
          ok: false,
          message:
            "“Only new signals” adds to the signals tab that is open, and the open tab is not one. Open the " +
            "signals tab to add to, or untick Only new signals to import into a new tab.",
        };
      }
      if (tab.rows > 0 && tab.newestCreatedAt == null) {
        return {
          ok: false,
          message:
            `“${tab.name}” has no Created at value this add-on can read, so it cannot tell which signals are ` +
            "new. Untick Only new signals to import into a new tab.",
        };
      }
      const bound = signalsReadBound(tab.newestCreatedAt);
      if (bound) {
        filters.created_after = bound;
        storedOrder = true;
      }
    }

    const sheet = sheetWriter(SIGNAL_COLUMNS, "Jobo Signals", tab ? tab.sheet : null);
    // Nothing in the tab is asked for again, so this set should never match.
    // It stays as the last guard against writing a signal twice, and the
    // result message says so if it ever does.
    const known = new Set(tab ? tab.ids : []);
    const chunks: string[][] = [];
    for (let i = 0; i < domains.length; i += SIGNAL_MAX_LIST_VALUES) chunks.push(domains.slice(i, i + SIGNAL_MAX_LIST_VALUES));
    if (chunks.length === 0) chunks.push([]);

    const usage = newUsage();
    let read = 0;
    let skipped = 0;
    let requests = 0;
    let failure: string | null = null;
    let outOfTime = false;
    let moreMatched = false;
    let unreadDomains = 0;

    reading: for (let c = 0; c < chunks.length; c++) {
      let cursor: string | undefined;
      for (;;) {
        if (read >= limit) {
          // Stopped by the cap: either this chunk had another page, or
          // chunks of domains are left that were never asked about.
          if (cursor) moreMatched = true;
          else for (let rest = c; rest < chunks.length; rest++) unreadDomains += chunks[rest].length;
          break reading;
        }
        if (requests > 0 && Date.now() >= deadline) {
          outOfTime = true;
          break reading;
        }

        // The order goes on every page: a cursor continues the order it was issued under.
        const query = toSignalsQuery(
          { ...filters, domains: chunks[c].length > 0 ? chunks[c] : undefined },
          { limit: Math.min(SIGNAL_MAX_LIMIT, limit - read), cursor, ...(storedOrder ? { sort: SIGNAL_POLL_SORT } : {}) },
        );
        requests++;
        // Billed per signal returned, so never replayed after a lost answer.
        let result: ApiResult;
        try {
          result = request(`/api/signals${buildQuery(query)}`, key, undefined, { replaySafe: false, deadline });
        } catch (err) {
          failure = (err as Error).message;
          break reading;
        }
        if (result.status < 200 || result.status >= 300) {
          failure = describeSignalsFailure(result.status, result.body);
          break reading;
        }
        addUsage(usage, result);

        const page = (result.body ?? {}) as Partial<HiringSignalListResponse>;
        const batch = Array.isArray(page.signals) ? page.signals : [];
        read += batch.length;
        const fresh = batch.filter((signal) => signal && !known.has(signal.id));
        skipped += batch.length - fresh.length;
        for (const signal of fresh) known.add(signal.id);
        try {
          sheet.append(fresh);
        } catch (err) {
          failure = `Writing to the sheet failed: ${(err as Error).message}`;
          break reading;
        }

        cursor = page.has_more && page.next_cursor ? page.next_cursor : undefined;
        if (!cursor || batch.length === 0) break;
      }
    }

    const written = sheet.written();
    const spent = signalsUsageText(read, usage);
    const overlapNote =
      skipped > 0
        ? ` ${counted(skipped, "signal")} ${skipped === 1 ? "was" : "were"} already in the tab and left out, ` +
          `and ${skipped === 1 ? "is" : "are"} billed again.`
        : "";

    if (written === 0) {
      if (failure != null) return { ok: false, message: read > 0 ? `${failure}${overlapNote}${spent}` : failure };
      if (read > 0) return { ok: true, message: `No new signals.${overlapNote}${spent}` };
      return {
        ok: true,
        message: tab
          ? `No new signals since the last import into “${tab.name}”. Nothing was billed.`
          : "No signals matched those filters. Nothing was billed.",
      };
    }

    sheet.finish();

    const head = tab
      ? `Added ${counted(written, "new signal")} to “${sheet.name()}”.`
      : `Imported ${counted(written, "signal")} into “${sheet.name()}”.`;
    let stopped = "";
    if (failure != null) stopped = ` The import stopped early: ${failure}`;
    else if (outOfTime) stopped = " The import stopped early to finish inside Google's 6-minute limit for one run.";
    else if (moreMatched) stopped = ` More signals matched than the ${limit} asked for.`;
    else if (unreadDomains > 0) {
      stopped = ` The ${limit} signals asked for were reached before ${counted(unreadDomains, "domain")} could be read.`;
    }
    // What an import that stopped early leaves behind, and whether a later
    // "only new" import goes back for it.
    //  - Stored order, one chunk of domains: the rows written are everything
    //    stored up to the last one, so the next "only new" import carries on
    //    exactly there.
    //  - Otherwise it does not. A new tab is read newest detection day first,
    //    and several chunks of domains are read one after another, so what
    //    was not reached can be older than the newest signal written, and
    //    "only new" reads only after that one.
    const resumable = storedOrder && chunks.length === 1;
    const incomplete = !stopped
      ? ""
      : resumable
        ? " Signals are read in the order Jobo stored them, so the next “Only new signals” import continues " +
          "after the last one added."
        : " A later “Only new signals” import reads only signals stored after the newest one in the tab, so it " +
          "will not pick up the ones this import did not reach. To get them, import again into a new tab with a " +
          "higher How many signals or narrower filters.";
    return {
      ok: failure == null && !outOfTime,
      message: `${head}${stopped}${incomplete}${overlapNote}${spent}`,
    };
  } catch (err) {
    return { ok: false, message: (err as Error).message };
  }
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
globals.onInstall = onInstall;
globals.showSidebar = showSidebar;
globals.showSignalsSidebar = showSignalsSidebar;
globals.getSettings = getSettings;
globals.saveApiKey = saveApiKey;
globals.getConnectionStatus = getConnectionStatus;
globals.disconnect = disconnect;
// The name a sidebar opened before this version still calls.
globals.clearApiKey = disconnect;
globals.startSignIn = startSignIn;
// Called by Apps Script at /usercallback; the name is OAUTH_CALLBACK_FUNCTION.
globals.authCallback = authCallback;
globals.getFilterOptions = getFilterOptions;
globals.suggestLocations = suggestLocations;
globals.suggestCompanies = suggestCompanies;
globals.getOccupations = getOccupations;
globals.importJobs = importJobs;
globals.getSelectedDomains = getSelectedDomains;
globals.getSignalsTabStatus = getSignalsTabStatus;
globals.importSignals = importSignals;
