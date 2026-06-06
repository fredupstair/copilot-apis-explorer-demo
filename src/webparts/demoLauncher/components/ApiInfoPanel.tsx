// =============================================================================
// ApiInfoPanel.tsx
// -----------------------------------------------------------------------------
// "What is this API?" section, rendered at the top of every tab. Shows the HTTP
// method badge, the live-editable endpoint URL, the required-permissions table
// and a link to the official Microsoft Learn documentation.
// =============================================================================

import * as React from 'react';
import { Icon, Link, Stack, Text, TextField } from '@fluentui/react';
import { IApiDefinition } from './types';

export interface IApiInfoPanelProps {
  api: IApiDefinition;
  /** The fully-resolved, editable endpoint URL (kept in the parent's state). */
  endpointUrl: string;
  onEndpointUrlChange: (url: string) => void;
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

const ApiInfoPanel: React.FC<IApiInfoPanelProps> = ({
  api,
  endpointUrl,
  onEndpointUrlChange
}) => {
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

        {/* Required permissions table */}
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
              <td style={cellStyle}>{api.permissions.delegated ?? '— not supported —'}</td>
              <td style={cellStyle}>{api.permissions.application ?? '— not supported —'}</td>
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
