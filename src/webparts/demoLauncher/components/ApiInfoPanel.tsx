// =============================================================================
// ApiInfoPanel.tsx
// -----------------------------------------------------------------------------
// "What is this API?" section, rendered at the top of every tab. Shows the HTTP
// method badge, the live-editable endpoint URL, the required-permissions table
// and a link to the official Microsoft Learn documentation.
// =============================================================================

import * as React from 'react';
import { Icon, Link, Stack, Text, TextField } from '@fluentui/react';
import { AuthMode, IApiDefinition, IPermissionCheck } from './types';

export interface IApiInfoPanelProps {
  api: IApiDefinition;
  /** The fully-resolved, editable endpoint URL (kept in the parent's state). */
  endpointUrl: string;
  onEndpointUrlChange: (url: string) => void;
  /** Current auth mode – decides which permission column carries the live check. */
  authMode: AuthMode;
  /** Result of validating the active token against the API's required scopes. */
  permissionCheck: IPermissionCheck;
}

/** Shared table-cell style for the permissions table. */
const cellStyle: React.CSSProperties = {
  border: '1px solid #edebe9',
  padding: '6px 10px',
  verticalAlign: 'top'
};

/** Coloured pill for the HTTP verb (green GET / blue POST). */
const MethodBadge: React.FC<{ method: string }> = ({ method }) => (
  <span
    style={{
      background: method === 'POST' ? '#0078d4' : '#107c10',
      color: '#fff',
      borderRadius: 4,
      padding: '2px 10px',
      fontWeight: 700,
      fontSize: 13,
      letterSpacing: 0.5
    }}
  >
    {method}
  </span>
);

/** Editable URL with syntax-highlighted dark overlay. A transparent textarea
 *  sits on top of a pretty-printed <pre> so the user types directly into the
 *  same colored view (origin in blue, path in yellow, query params one per
 *  line in cyan/orange). Both layers share identical font / padding / wrap
 *  rules so the textarea caret aligns with the highlighted glyphs. */
const EditableUrl: React.FC<{ value: string; onChange: (v: string) => void }> = ({
  value,
  onChange
}) => {
  // Highlighted view uses '\n  ?' / '\n  &' separators for readability while
  // the raw textarea keeps the real, single-line URL.
  const renderHighlight = (url: string): React.ReactNode => {
    if (!url) {
      return <span style={{ color: '#808080' }}>(empty)</span>;
    }
    const qIdx = url.indexOf('?');
    const base = qIdx === -1 ? url : url.substring(0, qIdx);
    const query = qIdx === -1 ? '' : url.substring(qIdx + 1);
    const schemeMatch = /^([a-z][a-z0-9+.-]*:\/\/[^/]+)(\/.*)?$/i.exec(base);
    const origin = schemeMatch ? schemeMatch[1] : '';
    const path = schemeMatch ? schemeMatch[2] ?? '' : base;
    const params = query.length > 0 ? query.split('&') : [];
    return (
      <>
        {origin && <span style={{ color: '#569cd6' }}>{origin}</span>}
        {path && <span style={{ color: '#dcdcaa' }}>{path}</span>}
        {params.length > 0 && (
          <>
            <span style={{ color: '#d4d4d4' }}>?</span>
            {params.map((p, i) => {
              const eq = p.indexOf('=');
              const name = eq === -1 ? p : p.substring(0, eq);
              const val = eq === -1 ? undefined : p.substring(eq + 1);
              return (
                <React.Fragment key={i}>
                  {i > 0 && <span style={{ color: '#d4d4d4' }}>&</span>}
                  <span style={{ color: '#9cdcfe' }}>{name}</span>
                  {val !== undefined && (
                    <>
                      <span style={{ color: '#d4d4d4' }}>=</span>
                      <span style={{ color: '#ce9178' }}>{val}</span>
                    </>
                  )}
                </React.Fragment>
              );
            })}
          </>
        )}
        {/* Trailing space guarantees the <pre> grows by one line when the
            textarea ends with a newline (browsers ignore a trailing \n). */}
        {'\u200b'}
      </>
    );
  };

  const sharedStyle: React.CSSProperties = {
    margin: 0,
    padding: '10px 12px',
    fontFamily: 'Consolas, "Courier New", monospace',
    fontSize: 12,
    lineHeight: 1.5,
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-all',
    border: '1px solid #3c3c3c',
    borderRadius: 3,
    boxSizing: 'border-box'
  };

  return (
    <div style={{ position: 'relative', background: '#1e1e1e', borderRadius: 3 }}>
      <pre
        aria-hidden
        style={{
          ...sharedStyle,
          color: '#d4d4d4',
          background: 'transparent',
          minHeight: '2.5em',
          pointerEvents: 'none'
        }}
      >
        {renderHighlight(value)}
      </pre>
      <textarea
        aria-label="Endpoint URL"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        spellCheck={false}
        wrap="soft"
        style={{
          ...sharedStyle,
          position: 'absolute',
          top: 0,
          left: 0,
          width: '100%',
          height: '100%',
          background: 'transparent',
          color: 'transparent',
          caretColor: '#ffffff',
          resize: 'none',
          outline: 'none',
          overflow: 'hidden',
          // Make selection visible against the dark bg.
          // (Firefox honors the textarea's selection color directly.)
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          ...({ WebkitTextFillColor: 'transparent' } as any)
        }}
      />
    </div>
  );
};

/** Render a single scope cell. When `active` is true the chip is coloured
 *  green / red based on whether the current token carries that scope; when
 *  `active` is false the chip is shown in a neutral grey (the API supports
 *  the column but it is not the auth mode we're currently using). */
const ScopeCell: React.FC<{
  scope: string | undefined;
  active: boolean;
  granted: string[];
}> = ({ scope, active, granted }) => {
  if (!scope) {
    return <span style={{ color: '#a19f9d' }}>— not supported —</span>;
  }
  if (!active) {
    return (
      <span
        style={{
          display: 'inline-flex',
          alignItems: 'center',
          gap: 4,
          padding: '2px 10px',
          borderRadius: 12,
          fontSize: 12,
          fontWeight: 600,
          color: '#605e5c',
          background: '#f3f2f1',
          border: '1px solid #edebe9'
        }}
        title="Switch auth mode to validate this scope against the active token."
      >
        {scope}
      </span>
    );
  }
  // Scope strings can list multiple required permissions separated by commas
  // (e.g. "Files.Read.All, Sites.Read.All"). Render one chip per individual
  // scope so each one is validated against the token independently.
  const parts = scope.split(',').map((s) => s.trim()).filter((s) => s.length > 0);
  return (
    <span style={{ display: 'inline-flex', flexWrap: 'wrap', gap: 4 }}>
      {parts.map((p) => {
        const ok = granted.indexOf(p) !== -1;
        return (
          <span
            key={p}
            style={{
              display: 'inline-flex',
              alignItems: 'center',
              gap: 4,
              padding: '2px 10px',
              borderRadius: 12,
              fontSize: 12,
              fontWeight: 600,
              color: '#ffffff',
              background: ok ? '#107c10' : '#a4262c'
            }}
            title={ok ? 'Scope present in the active token.' : 'Scope missing from the active token.'}
          >
            <Icon iconName={ok ? 'CheckMark' : 'Cancel'} /> {p}
          </span>
        );
      })}
    </span>
  );
};

const ApiInfoPanel: React.FC<IApiInfoPanelProps> = ({
  api,
  endpointUrl,
  onEndpointUrlChange,
  authMode,
  permissionCheck
}) => {
  const delegatedActive = authMode === 'graphClient' && !!api.permissions.delegated;
  const applicationActive = authMode === 'manualToken' && !!api.permissions.application;
  return (
    <div
      style={{
        border: '1px solid #edebe9',
        borderRadius: 4,
        padding: 16
      }}
    >
      <Stack tokens={{ childrenGap: 12 }}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
          <MethodBadge method={api.method} />
          <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
            {api.title}
          </Text>
        </Stack>

        <Stack tokens={{ childrenGap: 4 }}>
          <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
            Endpoint URL (editable – path &amp; query params are live)
          </Text>
          <EditableUrl value={endpointUrl} onChange={onEndpointUrlChange} />
        </Stack>

        {/* Required permissions table – the cell for the active auth mode
            carries the live green / red permission check. */}
        <table
          style={{
            borderCollapse: 'collapse',
            width: '100%',
            fontSize: 13
          }}
        >
          <thead>
            <tr style={{ background: '#f3f2f1', textAlign: 'left' }}>
              <th style={cellStyle}>Delegated scope</th>
              <th style={cellStyle}>Application scope</th>
              <th style={cellStyle}>Admin consent</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td style={cellStyle}>
                <ScopeCell
                  scope={api.permissions.delegated}
                  active={delegatedActive}
                  granted={permissionCheck.granted}
                />
              </td>
              <td style={cellStyle}>
                <ScopeCell
                  scope={api.permissions.application}
                  active={applicationActive}
                  granted={permissionCheck.granted}
                />
              </td>
              <td style={cellStyle}>
                {api.permissions.adminConsent ? (
                  <span style={{ color: '#a4262c', fontWeight: 600 }}>
                    <Icon iconName="Shield" /> Required
                  </span>
                ) : (
                  <span style={{ color: '#107c10', fontWeight: 600 }}>Not required</span>
                )}
              </td>
            </tr>
          </tbody>
        </table>

        <Link href={api.docsUrl} target="_blank" rel="noreferrer">
          <Icon iconName="NavigateExternalInline" /> Official Microsoft Learn documentation
        </Link>
      </Stack>
    </div>
  );
};

export default ApiInfoPanel;
