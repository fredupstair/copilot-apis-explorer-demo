// =============================================================================
// types.ts
// -----------------------------------------------------------------------------
// Single source of truth for every request/response shape and auth structure
// used by the Copilot API Explorer.
//
// Architectural decision (for the conference audience):
//   The whole web part is data-driven. Instead of hard-coding five different
//   tabs, we describe each API once with an `IApiDefinition`. Every UI piece
//   (info panel, request builder, response viewer, permission matrix) simply
//   reads that definition. Adding a 6th API later is therefore a *data* change,
//   not a *code* change.
// =============================================================================

/** Stable identifier for each of the five demonstrated APIs. */
export type ApiId =
  | 'retrieval'
  | 'interactionExport'
  | 'usageSummary'
  | 'usageTrend'
  | 'usageUserDetail';

/** HTTP verbs we actually use in this demo. */
export type HttpMethod = 'GET' | 'POST';

/**
 * How the call is authenticated.
 *  - graphClient: delegated auth via MSGraphClientV3 (current signed-in user).
 *  - manualToken: a bearer token pasted by the presenter (App Registration /
 *    client-credentials flow) used in a raw fetch(). DEMO ONLY.
 */
export type AuthMode = 'graphClient' | 'manualToken';

/** A single editable query-string parameter for GET APIs. */
export interface IQueryParam {
  /** Parameter key as it appears in the URL (e.g. "$top", "period", "userId"). */
  key: string;
  /** Current value (string – the UI keeps everything editable as text). */
  value: string;
  /**
   * Where the parameter belongs:
   *  - 'query'    => appended as ?key=value
   *  - 'odataFn'  => OData function argument, e.g. getMicrosoft365CopilotUserCountTrend(period='D7')
   *  - 'path'     => replaces a {placeholder} token in the URL template
   */
  location: 'query' | 'odataFn' | 'path';
  /** Short human description shown next to the field. */
  hint?: string;
}

/** The permission requirements for one API, rendered in info panel + matrix. */
export interface IApiPermissions {
  /** Delegated scope(s) – null when the API does not support delegated auth. */
  delegated: string | undefined;
  /** Application role(s) – null when the API does not support app-only auth. */
  application: string | undefined;
  /** Whether a tenant admin must grant consent (true for all .All scopes). */
  adminConsent: boolean;
}

/** Full, declarative description of one API. */
export interface IApiDefinition {
  id: ApiId;
  /** Short label used on the Pivot tab. */
  tabLabel: string;
  /** Friendly title shown in the info panel. */
  title: string;
  method: HttpMethod;
  /**
   * Endpoint template. May contain a {userId} placeholder (interaction export)
   * and/or an OData function segment. The presenter can still edit the final
   * URL live in the UI.
   */
  urlTemplate: string;
  /** Microsoft Learn documentation link. */
  docsUrl: string;
  permissions: IApiPermissions;
  /** True when MSGraphClientV3 (delegated) can be used for this API. */
  supportsDelegated: boolean;
  /** True when only an application token (manual paste) works. */
  supportsApp: boolean;
  /** Pre-filled JSON body for POST APIs (Retrieval only). */
  defaultBody?: IRetrievalRequest;
  /** Pre-filled editable query parameters for GET APIs. */
  defaultParams?: IQueryParam[];
  /** One-line explanation surfaced in Demo Mode. */
  demoNote: string;
}

// -----------------------------------------------------------------------------
// 1. Retrieval API  – POST /copilot/retrieval
// -----------------------------------------------------------------------------

/** Acceptable data sources for the Retrieval API. */
export type RetrievalDataSource = 'sharePoint' | 'oneDriveBusiness' | 'externalItem';

/** Request body for POST /copilot/retrieval (beta + v1.0). */
export interface IRetrievalRequest {
  queryString: string;
  dataSource: RetrievalDataSource;
  /** KQL scoping expression. Real API parameter name is `filterExpression`. */
  filterExpression?: string;
  resourceMetadata?: string[];
  /** 1–25. Kept as number; serialized as-is. */
  maximumNumberOfResults?: number;
}

export interface IRetrievalExtract {
  text: string;
  relevanceScore?: number;
}

export interface IRetrievalHit {
  webUrl: string;
  extracts: IRetrievalExtract[];
  resourceType?: string;
  resourceMetadata?: Record<string, string>;
  sensitivityLabel?: {
    displayName?: string;
    color?: string;
  };
}

export interface IRetrievalResponse {
  retrievalHits: IRetrievalHit[];
}

// -----------------------------------------------------------------------------
// 2. Interaction Export – GET /copilot/users/{id}/interactionHistory/getAllEnterpriseInteractions
// -----------------------------------------------------------------------------

export interface IAiInteraction {
  id: string;
  sessionId?: string;
  requestId?: string;
  appClass?: string;
  interactionType?: string;
  conversationType?: string;
  createdDateTime?: string;
  locale?: string;
  body?: {
    contentType?: string;
    content?: string;
  };
  from?: {
    user?: { displayName?: string; id?: string };
    application?: { displayName?: string; id?: string };
  };
}

export interface IInteractionExportResponse {
  value: IAiInteraction[];
  '@odata.nextLink'?: string;
}

// -----------------------------------------------------------------------------
// 3-5. Usage reports – /reports/getCopilot*  (return generic row collections)
// -----------------------------------------------------------------------------

/**
 * The Copilot usage report endpoints return a `value` array of flat objects.
 * Because the exact column set evolves, we type rows loosely as string maps and
 * let the bar-chart renderer pick numeric columns dynamically.
 */
export interface IUsageReportRow {
  [column: string]: string | number | undefined;
}

export interface IUsageReportResponse {
  value: IUsageReportRow[];
}

// -----------------------------------------------------------------------------
// Auth – decoded JWT + permission check
// -----------------------------------------------------------------------------

/** Subset of JWT claims we decode and display client-side (no external lib). */
export interface IDecodedToken {
  /** User principal name / preferred_username / unique_name, whichever exists. */
  upn?: string;
  /** Delegated scopes (from the space-delimited `scp` claim). */
  scopes: string[];
  /** Application roles (from the `roles` claim) – present on app-only tokens. */
  roles: string[];
  /** Token audience (should be https://graph.microsoft.com). */
  audience?: string;
  /** Tenant id. */
  tenantId?: string;
  /** Expiry as a unix epoch (seconds). */
  exp?: number;
  /** True when this is an application token (roles present, no scp). */
  isAppToken: boolean;
  /** Raw decoded payload for the "show the JWT" demo moment. */
  raw: Record<string, unknown>;
}

/** Result of checking whether the current token satisfies an API's scopes. */
export interface IPermissionCheck {
  /** Scope/role names that are required for the selected API + auth mode. */
  required: string[];
  /** Scope/role names that are present in the current token. */
  granted: string[];
  /** Scope/role names that are required but missing. */
  missing: string[];
  /** True when every required scope is present. */
  satisfied: boolean;
}

// -----------------------------------------------------------------------------
// API call result – everything ResponseViewer needs to render
// -----------------------------------------------------------------------------

export interface IApiCallResult {
  /** HTTP status code (0 when the request never left the browser). */
  status: number;
  /** True for 2xx responses. */
  ok: boolean;
  /** Round-trip duration in milliseconds. */
  timeMs: number;
  /** Approximate payload size in kilobytes. */
  sizeKb: number;
  /** Parsed JSON body (or raw text wrapper when parsing fails). */
  body: unknown;
  /** Populated when the call threw before/around the response. */
  error?: string;
}
