// =============================================================================
// DemoLauncher.tsx  (root component – "Copilot API Explorer")
// -----------------------------------------------------------------------------
// Postman-style layout for a full-bleed SharePoint section:
//
//   ┌──────────────┬──────────────────────────────────────────────────────────┐
//   │              │  Pivot tabs (one per API)                                │
//   │              │  ──────────────────────────────────────────────────────  │
//   │   AuthPanel  │  ApiInfoPanel                                            │
//   │   (sidebar)  │  RequestBuilder                                          │
//   │              │  ResponseViewer (raw JSON + per-API card view)           │
//   │              │  PermissionMatrix                                        │
//   └──────────────┴──────────────────────────────────────────────────────────┘
//
// Architectural decisions worth calling out on stage:
//   * State is keyed by ApiId so every tab keeps its own request/response.
//   * The endpoint URL is a derived-but-editable value: editing params rebuilds
//     it, but power users can still hand-edit the final URL.
//   * Two execution paths share one result shape (IApiCallResult): delegated
//     (MSGraphClientV3) and app-only (raw fetch with a token the user pastes
//     into the auth sidebar after generating it out-of-band).
// =============================================================================

import * as React from 'react';
import {
  Icon,
  Pivot,
  PivotItem,
  Stack,
  Text,
  ThemeProvider
} from '@fluentui/react';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import type { IDemoLauncherProps } from './IDemoLauncherProps';
import { ApiId, AuthMode, IApiCallResult, IQueryParam, IRetrievalRequest } from './types';
import {
  API_CATALOG,
  API_ORDER,
  buildUrl,
  checkPermissions,
  decodeJwt,
  executeWithBearerToken,
  executeWithGraphClient
} from './copilotApi';
import AuthPanel from './AuthPanel';
import ApiInfoPanel from './ApiInfoPanel';
import RequestBuilder from './RequestBuilder';
import ResponseViewer from './ResponseViewer';
import PermissionMatrix from './PermissionMatrix';

// ---- helpers to build the initial, per-API state -----------------------------

/** Clone the default params so edits never mutate the shared catalog. */
function initialParams(): Record<ApiId, IQueryParam[]> {
  const record = {} as Record<ApiId, IQueryParam[]>;
  for (const id of API_ORDER) {
    record[id] = (API_CATALOG[id].defaultParams ?? []).map((p) => ({ ...p }));
  }
  return record;
}

/** Compute the initial endpoint URL for every API from its template + params. */
function initialEndpoints(p: Record<ApiId, IQueryParam[]>): Record<ApiId, string> {
  const record = {} as Record<ApiId, string>;
  for (const id of API_ORDER) {
    record[id] = buildUrl(API_CATALOG[id].urlTemplate, p[id]);
  }
  return record;
}

const DemoLauncher: React.FC<IDemoLauncherProps> = (props) => {
  const {
    context,
    userDisplayName,
    description
  } = props;

  // ---- top-level UI state ----
  const [activeApiId, setActiveApiId] = React.useState<ApiId>('retrieval');

  // ---- auth state ----
  const [authMode, setAuthMode] = React.useState<AuthMode>('graphClient');
  const [delegatedToken, setDelegatedToken] = React.useState<string>('');
  const [acquiringDelegated, setAcquiringDelegated] = React.useState<boolean>(false);
  const [appToken, setAppToken] = React.useState<string>('');

  // ---- per-API request state ----
  const [params, setParams] = React.useState<Record<ApiId, IQueryParam[]>>(initialParams);
  const [endpoints, setEndpoints] = React.useState<Record<ApiId, string>>(() =>
    initialEndpoints(initialParams())
  );
  const [bodyText, setBodyText] = React.useState<string>(
    JSON.stringify(API_CATALOG.retrieval.defaultBody, undefined, 2)
  );
  const [kqlFilter, setKqlFilter] = React.useState<string>('');

  // ---- per-API response state ----
  const [results, setResults] = React.useState<Partial<Record<ApiId, IApiCallResult>>>({});
  const [executingApi, setExecutingApi] = React.useState<ApiId | undefined>(undefined);

  // ---- derived values ----
  const api = API_CATALOG[activeApiId];
  const activeToken = authMode === 'manualToken' ? appToken : delegatedToken;
  const decoded = React.useMemo(() => decodeJwt(activeToken), [activeToken]);
  const permissionCheck = React.useMemo(
    () => checkPermissions(api, authMode, decoded),
    [api, authMode, decoded]
  );

  /** Retrieval body actually sent: bodyText parsed + the KQL filter injected. */
  const effectiveRetrievalBody = React.useMemo<IRetrievalRequest | undefined>(() => {
    if (api.method !== 'POST') {
      return undefined;
    }
    try {
      const parsed = JSON.parse(bodyText) as IRetrievalRequest;
      if (kqlFilter.trim().length > 0) {
        parsed.filterExpression = kqlFilter.trim();
      } else {
        delete parsed.filterExpression;
      }
      return parsed;
    } catch {
      return undefined; // invalid JSON – execution will report it
    }
  }, [api.method, bodyText, kqlFilter]);

  // ---- handlers ----------------------------------------------------------------

  const handleAcquireDelegatedToken = (): void => {
    setAcquiringDelegated(true);
    context.aadTokenProviderFactory
      .getTokenProvider()
      .then((provider) => provider.getToken('https://graph.microsoft.com'))
      .then((token) => {
        setDelegatedToken(token);
        setAuthMode('graphClient');
      })
      .catch((err: Error) => {
        setDelegatedToken('');
        // Surface the failure in the response area of the active tab.
        setResults((prev) => ({
          ...prev,
          [activeApiId]: {
            status: 0,
            ok: false,
            timeMs: 0,
            sizeKb: 0,
            body: { error: err.message },
            error: `Token acquisition failed: ${err.message}`
          }
        }));
      })
      .finally(() => setAcquiringDelegated(false));
  };

  const handleAppTokenChange = (token: string): void => {
    setAppToken(token);
    if (token.trim().length > 0) {
      setAuthMode('manualToken');
    }
  };

  const handleParamChange = (index: number, value: string): void => {
    setParams((prev) => {
      const next = { ...prev };
      const updated = next[activeApiId].map((p, i) => (i === index ? { ...p, value } : p));
      next[activeApiId] = updated;
      // Rebuild the editable endpoint URL so the preview stays in sync.
      setEndpoints((eps) => ({
        ...eps,
        [activeApiId]: buildUrl(API_CATALOG[activeApiId].urlTemplate, updated)
      }));
      return next;
    });
  };

  const handleEndpointUrlChange = (url: string): void => {
    setEndpoints((prev) => ({ ...prev, [activeApiId]: url }));
  };

  const handleExecute = (): void => {
    const url = endpoints[activeApiId];
    const method = api.method;
    const body = effectiveRetrievalBody;

    // Guard: invalid JSON body for the POST API.
    if (method === 'POST' && !body) {
      setResults((prev) => ({
        ...prev,
        [activeApiId]: {
          status: 0,
          ok: false,
          timeMs: 0,
          sizeKb: 0,
          body: { error: 'Request body is not valid JSON.' },
          error: 'Request body is not valid JSON.'
        }
      }));
      return;
    }

    setExecutingApi(activeApiId);

    const finish = (result: IApiCallResult): void => {
      setResults((prev) => ({ ...prev, [activeApiId]: result }));
      setExecutingApi(undefined);
    };

    if (authMode === 'graphClient') {
      context.msGraphClientFactory
        .getClient('3')
        .then((client: MSGraphClientV3) => executeWithGraphClient(client, method, url, body))
        .then(finish)
        .catch((err: Error) =>
          finish({
            status: 0,
            ok: false,
            timeMs: 0,
            sizeKb: 0,
            body: { error: err.message },
            error: err.message
          })
        );
    } else {
      executeWithBearerToken(appToken, method, url, body)
        .then(finish)
        .catch((err: Error) =>
          finish({
            status: 0,
            ok: false,
            timeMs: 0,
            sizeKb: 0,
            body: { error: err.message },
            error: err.message
          })
        );
    }
  };

  // ---- execute gating ----------------------------------------------------------

  let canExecute = true;
  let disabledReason: string | undefined;
  if (authMode === 'graphClient' && !api.supportsDelegated) {
    canExecute = false;
    disabledReason =
      'This API is Application-only. Switch to "App Registration (Application)" and paste an app-only token (Reports.Read.All / AiEnterpriseInteraction.Read.All).';
  } else if (authMode === 'manualToken' && appToken.trim().length === 0) {
    canExecute = false;
    disabledReason = 'Paste an app-only token in the Authentication panel to execute.';
  } else if (authMode === 'manualToken' && !api.supportsApp) {
    canExecute = false;
    disabledReason =
      'This API does not support Application permissions. Switch to "Delegated" mode.';
  }

  // ---- render ------------------------------------------------------------------

  return (
    <ThemeProvider>
      <Stack tokens={{ childrenGap: 16 }} styles={{ root: { padding: 16 } }}>
        {/* Header */}
        <Stack>
          <Text variant="xLarge" styles={{ root: { fontWeight: 700 } }}>
            <Icon iconName="Robot" /> Copilot API Explorer
          </Text>
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            {description ? `${description} · ` : ''}Signed in as {userDisplayName}
          </Text>
        </Stack>

        {/* Two-column workspace: sticky auth sidebar + main Postman pane */}
        <div
          style={{
            display: 'grid',
            gridTemplateColumns: 'minmax(320px, 360px) minmax(0, 1fr)',
            gap: 16,
            alignItems: 'start'
          }}
        >
          {/* --- Left column: authentication (sticky) --- */}
          <div style={{ position: 'sticky', top: 8 }}>
            <AuthPanel
              authMode={authMode}
              onAuthModeChange={setAuthMode}
              delegatedToken={delegatedToken}
              onAcquireDelegatedToken={handleAcquireDelegatedToken}
              acquiringDelegated={acquiringDelegated}
              appToken={appToken}
              onAppTokenChange={handleAppTokenChange}
              decoded={decoded}
              permissionCheck={permissionCheck}
              selectedApiTitle={api.title}
            />
          </div>

          {/* --- Right column: Postman-style tabs + per-API workspace --- */}
          <div
            style={{
              border: '1px solid #edebe9',
              borderRadius: 4,
              background: '#ffffff',
              padding: 12,
              minWidth: 0
            }}
          >
            <Pivot
              selectedKey={activeApiId}
              onLinkClick={(item) => {
                if (item?.props.itemKey) {
                  setActiveApiId(item.props.itemKey as ApiId);
                }
              }}
              styles={{ root: { borderBottom: '1px solid #edebe9', marginBottom: 12 } }}
            >
              {API_ORDER.map((id) => (
                <PivotItem
                  key={id}
                  itemKey={id}
                  headerText={API_CATALOG[id].tabLabel}
                />
              ))}
            </Pivot>

            <Stack tokens={{ childrenGap: 16 }}>
              <ApiInfoPanel
                api={api}
                endpointUrl={endpoints[activeApiId]}
                onEndpointUrlChange={handleEndpointUrlChange}
              />
              <RequestBuilder
                api={api}
                endpointUrl={endpoints[activeApiId]}
                activeToken={activeToken}
                bodyText={bodyText}
                onBodyTextChange={setBodyText}
                kqlFilter={kqlFilter}
                onKqlFilterChange={setKqlFilter}
                effectiveBody={effectiveRetrievalBody}
                params={params[activeApiId]}
                onParamChange={handleParamChange}
              />
              <ResponseViewer
                api={api}
                result={results[activeApiId]}
                executing={executingApi === activeApiId}
                onExecute={handleExecute}
                canExecute={canExecute}
                disabledReason={disabledReason}
              />
              <PermissionMatrix activeApiId={activeApiId} />
            </Stack>
          </div>
        </div>
      </Stack>
    </ThemeProvider>
  );
};

export default DemoLauncher;
