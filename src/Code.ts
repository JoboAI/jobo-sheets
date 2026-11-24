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
 *    Whether a failure is retried, and how long to wait, is still decided by
 *    connector-core (`isRetryable`, `nextDelayMs`), so the policy cannot drift.
 *  - **A run ends after 6 minutes.** An import writes each page as it arrives
 *    and stops starting new work after `IMPORT_TIME_BUDGET_MS`.
 */

import {
  COMPANY_CATEGORIES,
  EMPLOYMENT_TYPES,
  EXPERIENCE_LEVELS,
  FeedCursorRestartRequiredError,
  InsufficientCreditsError,
  JoboApiError,
  JoboTransportError,
  LANGUAGES,
  RateLimitError,
  WORK_MODELS,
  defaultRetryOptions,
  isRetryable,
  isValidApiKeyFormat,
  nextDelayMs,
  toJobSearchBody,
  type CompanySuggestion,
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
   * False for a request billed for what it returns. Default true: every other
   * endpoint this add-on calls is free.
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
//
// This is an editor add-on: a menu under Extensions and an HTML sidebar. The
// manifest therefore has no `addOns` section; that section, and its homepage
// trigger that must return Cards, belong to Google Workspace add-ons.
//
// onOpen and onInstall can run before the user has authorised the add-on in
// this file (AuthMode.NONE), so they only build the menu.

function onOpen(): void {
  SpreadsheetApp.getUi().createAddonMenu().addItem("Import jobs…", "showSidebar").addToUi();
}

/** Installing from the Marketplace does not run onOpen for the file already open. */
function onInstall(): void {
  onOpen();
}

function showSidebar(): void {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar").setTitle("Career Site Jobs");
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── Called from the sidebar ──────────────────────────────────────────────────

function getSettings(): {
  hasKey: boolean;
  /** The signed-in account, when the key came from Sign in with Jobo. */
  accountEmail: string | null;
  workModels: string[];
  employmentTypes: string[];
  experienceLevels: string[];
  companyCategories: string[];
  languages: Array<{ code: string; name: string }>;
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

    const sheet = importSheet();
    let page = 1;
    let creditsSpent = 0;
    let balance: number | null = null;
    let quotaLimit: number | null = null;
    let quotaRemaining: number | null = null;
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

      creditsSpent += result.creditsDeducted ?? 0;
      balance = result.creditsBalance ?? balance;
      quotaLimit = result.quotaLimit ?? quotaLimit;
      quotaRemaining = result.quotaRemaining ?? quotaRemaining;

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

    const cost = creditsSpent > 0 ? ` Cost: $${(creditsSpent / 1000).toFixed(2)} (${creditsSpent} credits).` : "";
    const left = balance != null ? ` Wallet balance: $${(balance / 1000).toFixed(2)}.` : "";
    const included = quotaRemaining != null
      ? ` Included jobs: ${quotaRemaining}${quotaLimit != null ? ` of ${quotaLimit}` : ""} remaining.`
      : "";
    if (failure != null) {
      return {
        ok: false,
        message: `Imported ${written} of the ${limit} jobs requested before the import stopped. ${failure}${cost}${left}${included}${note}`,
      };
    }
    if (outOfTime) {
      return {
        ok: false,
        message:
          `Imported ${written} of the ${limit} jobs requested. The import stopped early to finish inside ` +
          `Google's 6-minute limit for one run. Importing again starts from the first matching job and bills ` +
          `again for the jobs already imported.${cost}${left}${included}${note}`,
      };
    }
    return { ok: true, message: `Imported ${written} jobs.${cost}${left}${included}${note}` };
  } catch (err) {
    return { ok: false, message: (err as Error).message };
  }
}

/**
 * The sheet one import writes to, created with its header row when the first
 * job arrives. Each append is one setValues call followed by a flush, so a page
 * is in the spreadsheet before the next one is requested: a run that is killed
 * or fails later keeps every page it was billed for.
 */
function importSheet(): { append: (jobs: Job[]) => void; finish: () => void; written: () => number } {
  let sheet: GoogleAppsScript.Spreadsheet.Sheet | null = null;
  let written = 0;

  return {
    written: () => written,
    append(jobs: Job[]): void {
      if (jobs.length === 0) return;
      const rows: Array<Array<string | number>> = [];
      // Header and first page in one write; formatted now, not at the end, so
      // a sheet left by a killed run still reads as a table.
      if (!sheet) rows.push(COLUMNS.map((c) => c.header));
      for (const job of jobs) rows.push(COLUMNS.map((c) => neutralizeFormula(c.value(job))));

      const firstRow = sheet ? written + 2 : 1;
      if (!sheet) {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        sheet = spreadsheet.insertSheet(uniqueSheetName(spreadsheet, "Jobo Jobs"));
        sheet.setFrozenRows(1);
      }
      sheet.getRange(firstRow, 1, rows.length, COLUMNS.length).setValues(rows);
      if (firstRow === 1) sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight("bold");
      SpreadsheetApp.flush();
      written += jobs.length;
    },
    finish(): void {
      // Once, at the end: resizing measures every row written.
      if (sheet) sheet.autoResizeColumns(1, COLUMNS.length);
    },
  };
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
