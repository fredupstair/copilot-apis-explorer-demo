// =============================================================================
// copilotApi.ts
// -----------------------------------------------------------------------------
// Pure, framework-agnostic logic for the Copilot API Explorer:
//   * API_CATALOG       – the declarative description of the five APIs.
//   * JWT decoding       – plain base64url, no external library.
//   * Permission checks  – does the current token carry the required scopes?
//   * URL building        – turn an IApiDefinition + params into a final URL.
//   * Request execution  – via MSGraphClientV3 (delegated) OR raw fetch (token).
//   * Exporters          – "Copy as cURL" and "Copy as Postman Collection".
//
// Keeping this here (instead of inside the React components) means the demo
// logic is unit-testable and the components stay presentational.
// =============================================================================

import { MSGraphClientV3 } from '@microsoft/sp-http';
import {
  ApiId,
  AuthMode,
  IApiCallResult,
  IApiDefinition,
  IDecodedToken,
  IPermissionCheck,
  IQueryParam
} from './types';

// -----------------------------------------------------------------------------
// The declarative catalog. Everything the UI knows about an API lives here.
// -----------------------------------------------------------------------------

export const API_CATALOG: Record<ApiId, IApiDefinition> = {
  retrieval: {
    id: 'retrieval',
    tabLabel: 'Retrieval',
    title: 'Retrieval API',
    method: 'POST',
    // Retrieval is delegated-only. v1.0 is the GA endpoint; the presenter can
    // still switch to /beta in the editable URL to demo newer behaviour.
    urlTemplate: 'https://graph.microsoft.com/v1.0/copilot/retrieval',
    docsUrl:
      'https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/retrieval/copilotroot-retrieval',
    permissions: {
      delegated: 'Files.Read.All, Sites.Read.All',
      application: undefined, // Not supported by the API.
      adminConsent: true
    },
    supportsDelegated: true,
    supportsApp: false,
    // NOTE: the real body uses queryString/dataSource/filterExpression – NOT the
    // entityTypes/from/size shape often seen in early drafts. We pre-fill the
    // schema that actually returns 200 OK.
    defaultBody: {
      queryString: 'Tell me more about Cowork',
      dataSource: 'sharePoint',
      resourceMetadata: ['title', 'author'],
      maximumNumberOfResults: 5
    },
    demoNote:
      'Grounded RAG over SharePoint/OneDrive/connectors. Results are permission-trimmed to the signed-in user.'
  },

  interactionExport: {
    id: 'interactionExport',
    tabLabel: 'Interaction Export',
    title: 'Interaction Export API',
    method: 'GET',
    urlTemplate:
      'https://graph.microsoft.com/beta/copilot/users/{userId}/interactionHistory/getAllEnterpriseInteractions',
    docsUrl:
      'https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/interaction-export/aiinteractionhistory-getallenterpriseinteractions',
    permissions: {
      delegated: undefined, // Not supported – application only.
      application: 'AiEnterpriseInteraction.Read.All',
      adminConsent: true
    },
    supportsDelegated: false,
    supportsApp: true,
    defaultParams: [
      {
        key: 'userId',
        value: '',
        location: 'path',
        hint: 'AAD object id of the target user (replaces {userId}).'
      },
      { key: '$top', value: '100', location: 'query', hint: 'Recommended page size is 100.' }
    ],
    demoNote:
      'Application-only export of every Copilot prompt/response for a user. MSGraphClient (delegated) cannot call this – use a manual app token.'
  },

  usageSummary: {
    id: 'usageSummary',
    tabLabel: 'Usage Summary',
    title: 'Copilot Usage – User Count Summary',
    method: 'GET',
    urlTemplate: "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUserCountSummary(period='{period}')",
    docsUrl:
      'https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/admin-settings/reports/resources/copilotreportroot',
    permissions: {
      delegated: 'Reports.Read.All',
      application: 'Reports.Read.All',
      adminConsent: true
    },
    supportsDelegated: true,
    supportsApp: true,
    defaultParams: [
      { key: 'period', value: 'D7', location: 'odataFn', hint: 'D7 | D30 | D90 | D180' }
    ],
    demoNote: 'Tenant-level adoption snapshot. Needs Reports.Read.All + admin consent.'
  },

  usageTrend: {
    id: 'usageTrend',
    tabLabel: 'Usage Trend',
    title: 'Copilot Usage – User Count Trend',
    method: 'GET',
    urlTemplate: "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUserCountTrend(period='{period}')",
    docsUrl:
      'https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/admin-settings/reports/resources/copilotreportroot',
    permissions: {
      delegated: 'Reports.Read.All',
      application: 'Reports.Read.All',
      adminConsent: true
    },
    supportsDelegated: true,
    supportsApp: true,
    defaultParams: [
      { key: 'period', value: 'D30', location: 'odataFn', hint: 'D7 | D30 | D90 | D180' }
    ],
    demoNote: 'Daily active-user trend over time – perfect for the CSS bar chart.'
  },

  usageUserDetail: {
    id: 'usageUserDetail',
    tabLabel: 'Usage User Detail',
    title: 'Copilot Usage – Per-User Detail',
    method: 'GET',
    urlTemplate: "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUsageUserDetail(period='{period}')",
    docsUrl:
      'https://learn.microsoft.com/en-us/graph/api/reportroot-getmicrosoft365copilotusageuserdetail?view=graph-rest-beta',
    permissions: {
      delegated: 'Reports.Read.All',
      application: 'Reports.Read.All',
      adminConsent: true
    },
    supportsDelegated: true,
    supportsApp: true,
    defaultParams: [
      { key: 'period', value: 'D7', location: 'odataFn', hint: 'D7 | D30 | D90 | D180' },
      { key: '$top', value: '20', location: 'query', hint: 'Limit rows for readability.' }
    ],
    demoNote: 'Per-user last-activity detail. Often returned as CSV; we request JSON.'
  }
};

/** Ordered list used to render the Pivot tabs deterministically. */
export const API_ORDER: ApiId[] = [
  'retrieval',
  'interactionExport',
  'usageSummary',
  'usageTrend',
  'usageUserDetail'
];

// -----------------------------------------------------------------------------
// Retrieval API – resourceMetadata fields the presenter can toggle on/off.
// The `apiName` is what we put inside the JSON body (`resourceMetadata: [...]`)
// while `label` is the friendly text shown next to the checkbox in the UI.
// -----------------------------------------------------------------------------
export interface IRetrievalField {
  label: string;
  apiName: string;
}

export const RETRIEVAL_FIELDS: IRetrievalField[] = [
  { label: 'File Name', apiName: 'fileName' },
  { label: 'Path', apiName: 'path' },
  { label: 'Author', apiName: 'author' },
  { label: 'Size', apiName: 'size' },
  { label: 'File Type', apiName: 'fileType' },
  { label: 'Title', apiName: 'title' },
  { label: 'Id', apiName: 'id' },
  { label: 'Drive Id', apiName: 'driveId' },
  { label: 'Site Id', apiName: 'siteId' },
  { label: 'List Id', apiName: 'listId' },
  { label: 'Created By', apiName: 'createdBy' },
  { label: 'Last Modified Time', apiName: 'lastModifiedDateTime' },
  { label: 'Modified By', apiName: 'lastModifiedBy' }
];

// -----------------------------------------------------------------------------
// JWT decoding – plain base64url, no library. Never trust this for security:
// the browser does NOT verify the signature. It is purely a demo aid.
// -----------------------------------------------------------------------------

function base64UrlDecode(segment: string): string {
  // Convert base64url -> base64 and pad to a multiple of 4.
  let base64 = segment.replace(/-/g, '+').replace(/_/g, '/');
  while (base64.length % 4 !== 0) {
    base64 += '=';
  }
  // atob handles Latin1; decodeURIComponent reconstructs UTF-8 characters.
  const binary = atob(base64);
  const percentEncoded = binary
    .split('')
    .map((char) => '%' + ('00' + char.charCodeAt(0).toString(16)).slice(-2))
    .join('');
  return decodeURIComponent(percentEncoded);
}

/**
 * Decode the payload section of a JWT. Returns undefined when the input is not a
 * well-formed three-part token, so the UI can show a friendly message.
 */
export function decodeJwt(token: string): IDecodedToken | undefined {
  const trimmed = token.trim();
  const parts = trimmed.split('.');
  if (parts.length !== 3) {
    return undefined;
  }

  try {
    const payload = JSON.parse(base64UrlDecode(parts[1])) as Record<string, unknown>;

    const scp = typeof payload.scp === 'string' ? payload.scp : '';
    const scopes = scp.length > 0 ? scp.split(' ').filter(Boolean) : [];
    const roles = Array.isArray(payload.roles) ? (payload.roles as string[]) : [];

    const upn =
      (payload.upn as string) ||
      (payload.preferred_username as string) ||
      (payload.unique_name as string) ||
      undefined;

    return {
      upn,
      scopes,
      roles,
      audience: payload.aud as string | undefined,
      tenantId: payload.tid as string | undefined,
      exp: typeof payload.exp === 'number' ? payload.exp : undefined,
      // App-only tokens carry roles and no delegated scopes.
      isAppToken: roles.length > 0 && scopes.length === 0,
      raw: payload
    };
  } catch {
    return undefined;
  }
}

/** Format the remaining lifetime of a token as a friendly countdown string. */
export function formatExpiryCountdown(exp: number | undefined): string {
  if (!exp) {
    return 'unknown';
  }
  const secondsLeft = exp - Math.floor(Date.now() / 1000);
  if (secondsLeft <= 0) {
    return 'EXPIRED';
  }
  const h = Math.floor(secondsLeft / 3600);
  const m = Math.floor((secondsLeft % 3600) / 60);
  const s = secondsLeft % 60;
  return `${h}h ${m}m ${s}s`;
}

// -----------------------------------------------------------------------------
// Permission checks
// -----------------------------------------------------------------------------

/** Split a "Foo.Read.All, Bar.Read.All" string into individual scope names. */
export function splitScopes(scopeString: string | undefined): string[] {
  if (!scopeString) {
    return [];
  }
  return scopeString
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
}

/**
 * Compare the scopes required by an API (for the active auth mode) against the
 * scopes/roles carried by the current token.
 */
export function checkPermissions(
  api: IApiDefinition,
  authMode: AuthMode,
  decoded: IDecodedToken | undefined
): IPermissionCheck {
  // Delegated mode validates against `scp`; app mode validates against `roles`.
  const required =
    authMode === 'manualToken' && decoded?.isAppToken
      ? splitScopes(api.permissions.application)
      : splitScopes(api.permissions.delegated);

  const held = decoded ? (decoded.isAppToken ? decoded.roles : decoded.scopes) : [];

  // Microsoft Graph scope comparison is case-insensitive.
  const heldLower = held.map((s) => s.toLowerCase());
  const granted = required.filter((r) => heldLower.indexOf(r.toLowerCase()) !== -1);
  const missing = required.filter((r) => heldLower.indexOf(r.toLowerCase()) === -1);

  return {
    required,
    granted,
    missing,
    satisfied: required.length > 0 && missing.length === 0
  };
}

// -----------------------------------------------------------------------------
// URL building from an API definition + live params
// -----------------------------------------------------------------------------

/**
 * Produce the final, fully-resolved URL for a GET/POST API.
 *  - {placeholder} path tokens are substituted.
 *  - OData function args (period) are injected inside the (...) segment.
 *  - remaining 'query' params are appended as a query string.
 */
export function buildUrl(template: string, params: IQueryParam[]): string {
  let url = template;

  // 1. Path placeholders, e.g. {userId}.
  for (const p of params.filter((x) => x.location === 'path')) {
    url = url.replace(`{${p.key}}`, encodeURIComponent(p.value));
  }

  // 2. OData function arguments, e.g. getMicrosoft365CopilotUserCountTrend(period='{period}').
  for (const p of params.filter((x) => x.location === 'odataFn')) {
    url = url.replace(`{${p.key}}`, p.value);
  }

  // 3. Classic query string params.
  const queryParams = params.filter((x) => x.location === 'query' && x.value.trim().length > 0);
  if (queryParams.length > 0) {
    const qs = queryParams
      .map((p) => `${encodeURIComponent(p.key)}=${encodeURIComponent(p.value)}`)
      .join('&');
    url += (url.indexOf('?') === -1 ? '?' : '&') + qs;
  }

  return url;
}

/**
 * Split a full Graph URL into the version segment ('v1.0' | 'beta') and the
 * relative path that MSGraphClientV3.api() expects.
 */
export function splitGraphUrl(fullUrl: string): { version: string; relativePath: string } {
  const match = /graph\.microsoft\.com\/(v1\.0|beta)(\/.*)$/i.exec(fullUrl);
  if (match) {
    return { version: match[1], relativePath: match[2] };
  }
  // Fallback: assume beta and treat everything after the host as the path.
  const hostStripped = fullUrl.replace(/^https?:\/\/[^/]+/i, '');
  return { version: 'beta', relativePath: hostStripped || '/' };
}

// -----------------------------------------------------------------------------
// Request execution
// -----------------------------------------------------------------------------

/** Approximate byte size (KB) of a JSON-serialisable value. */
function sizeInKb(value: unknown, fallbackText?: string): number {
  try {
    const text = fallbackText ?? JSON.stringify(value);
    return Math.round((new TextEncoder().encode(text).length / 1024) * 100) / 100;
  } catch {
    return 0;
  }
}

/**
 * Execute the call using MSGraphClientV3 (delegated, current user).
 * The factory is obtained by the caller; we receive a ready client here.
 */
export async function executeWithGraphClient(
  client: MSGraphClientV3,
  method: 'GET' | 'POST',
  fullUrl: string,
  body?: unknown
): Promise<IApiCallResult> {
  const { version, relativePath } = splitGraphUrl(fullUrl);
  const started = performance.now();

  try {
    const request = client.api(relativePath).version(version);
    const json: unknown =
      method === 'POST' ? await request.post(body) : await request.get();
    const timeMs = Math.round(performance.now() - started);

    return {
      status: 200,
      ok: true,
      timeMs,
      sizeKb: sizeInKb(json),
      body: json
    };
  } catch (err) {
    const timeMs = Math.round(performance.now() - started);
    // The Graph client surfaces statusCode + a JSON body string on failures.
    const e = err as { statusCode?: number; message?: string; body?: string };
    let parsedBody: unknown = e.body;
    if (typeof e.body === 'string') {
      try {
        parsedBody = JSON.parse(e.body);
      } catch {
        /* leave as string */
      }
    }
    return {
      status: e.statusCode ?? 0,
      ok: false,
      timeMs,
      sizeKb: sizeInKb(parsedBody, typeof e.body === 'string' ? e.body : undefined),
      body: parsedBody ?? { error: e.message },
      error: e.message
    };
  }
}

/**
 * Execute the call with a raw fetch() and a manually-pasted bearer token.
 * DEMO ONLY: a real app must never hold a raw access token in the browser.
 */
export async function executeWithBearerToken(
  token: string,
  method: 'GET' | 'POST',
  fullUrl: string,
  body?: unknown
): Promise<IApiCallResult> {
  const started = performance.now();
  try {
    const response = await fetch(fullUrl, {
      method,
      headers: {
        Authorization: `Bearer ${token.trim()}`,
        ...(method === 'POST' ? { 'Content-Type': 'application/json' } : {})
      },
      body: method === 'POST' ? JSON.stringify(body) : undefined
    });

    const text = await response.text();
    const timeMs = Math.round(performance.now() - started);

    let parsed: unknown = text;
    try {
      parsed = text.length > 0 ? JSON.parse(text) : {};
    } catch {
      /* non-JSON response (e.g. CSV) – keep raw text */
    }

    return {
      status: response.status,
      ok: response.ok,
      timeMs,
      sizeKb: sizeInKb(parsed, text),
      body: parsed,
      error: response.ok ? undefined : `HTTP ${response.status} ${response.statusText}`
    };
  } catch (err) {
    const timeMs = Math.round(performance.now() - started);
    const message = err instanceof Error ? err.message : String(err);
    return {
      status: 0,
      ok: false,
      timeMs,
      sizeKb: 0,
      body: { error: message },
      error: message
    };
  }
}

// -----------------------------------------------------------------------------
// Exporters: cURL + Postman v2.1
// -----------------------------------------------------------------------------

/** Build a copy-pasteable cURL command for the current request + token. */
export function buildCurl(
  method: 'GET' | 'POST',
  fullUrl: string,
  token: string,
  body?: unknown
): string {
  const lines: string[] = [`curl -X ${method} "${fullUrl}"`];
  lines.push(`  -H "Authorization: Bearer ${token.trim() || '<ACCESS_TOKEN>'}"`);
  if (method === 'POST') {
    lines.push(`  -H "Content-Type: application/json"`);
    const json = JSON.stringify(body ?? {}, undefined, 2).replace(/'/g, "'\\''");
    lines.push(`  -d '${json}'`);
  }
  // Backslash-newline continuation for readability when pasted into a shell.
  return lines.join(' \\\n');
}

/** Build a valid Postman v2.1 collection containing just this request. */
export function buildPostmanCollection(
  name: string,
  method: 'GET' | 'POST',
  fullUrl: string,
  token: string,
  body?: unknown
): string {
  const urlObj = new URL(fullUrl);
  const collection = {
    info: {
      name: `Copilot API Explorer – ${name}`,
      schema: 'https://schema.getpostman.com/json/collection/v2.1.0/collection.json'
    },
    item: [
      {
        name,
        request: {
          method,
          header: [
            { key: 'Authorization', value: `Bearer ${token.trim() || '{{access_token}}'}` },
            ...(method === 'POST'
              ? [{ key: 'Content-Type', value: 'application/json' }]
              : [])
          ],
          ...(method === 'POST'
            ? {
                body: {
                  mode: 'raw',
                  raw: JSON.stringify(body ?? {}, undefined, 2),
                  options: { raw: { language: 'json' } }
                }
              }
            : {}),
          url: {
            raw: fullUrl,
            protocol: urlObj.protocol.replace(':', ''),
            host: urlObj.hostname.split('.'),
            path: urlObj.pathname.split('/').filter(Boolean),
            query: Array.from(urlObj.searchParams.entries()).map(([key, value]) => ({
              key,
              value
            }))
          }
        }
      }
    ]
  };
  return JSON.stringify(collection, undefined, 2);
}

/** Small clipboard helper that degrades gracefully on older browsers. */
export async function copyToClipboard(text: string): Promise<void> {
  if (navigator.clipboard && navigator.clipboard.writeText) {
    await navigator.clipboard.writeText(text);
    return;
  }
  // Fallback for environments without the async clipboard API.
  const textarea = document.createElement('textarea');
  textarea.value = text;
  textarea.style.position = 'fixed';
  textarea.style.opacity = '0';
  document.body.appendChild(textarea);
  textarea.select();
  document.execCommand('copy');
  document.body.removeChild(textarea);
}

// -----------------------------------------------------------------------------
// OAuth 2.0 client_credentials – intentionally NOT implemented in the browser.
// The Microsoft identity platform does NOT expose CORS for the
// client_credentials grant on the v2.0 token endpoint, and a confidential
// client by definition cannot ship a client secret to JavaScript.
//
// The supported demo flow is therefore: generate the app-only token outside
// the browser (see auth/Get-AppToken.ps1) and paste it into the AuthPanel.
// -----------------------------------------------------------------------------
