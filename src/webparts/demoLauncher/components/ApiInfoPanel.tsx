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

        <TextField
          label="Endpoint URL (editable – path & query params are live)"
          value={endpointUrl}
          onChange={(_e, v) => onEndpointUrlChange(v ?? '')}
          styles={{ field: { fontFamily: 'Consolas, monospace', fontSize: 12 } }}
        />

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
