// =============================================================================
// PermissionMatrix.tsx
// -----------------------------------------------------------------------------
// A static, at-a-glance comparison of the permission requirements across all
// five APIs. Rendered at the bottom of every tab; the row for the currently
// active API is highlighted so the audience keeps their bearings.
// =============================================================================

import * as React from 'react';
import { Icon, Stack, Text } from '@fluentui/react';
import { ApiId } from './types';
import { API_CATALOG, API_ORDER } from './copilotApi';

export interface IPermissionMatrixProps {
  /** The API tab currently in focus – its row is highlighted. Omit to leave no row highlighted. */
  activeApiId?: ApiId;
}

const cellStyle: React.CSSProperties = {
  border: '1px solid #edebe9',
  padding: '6px 10px',
  fontSize: 12,
  verticalAlign: 'top'
};

const PermissionMatrix: React.FC<IPermissionMatrixProps> = ({ activeApiId }) => {
  return (
    <div style={{ border: '1px solid #edebe9', borderRadius: 4, padding: 16 }}>
      <Stack tokens={{ childrenGap: 10 }}>
        <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
          <Icon iconName="Table" /> Permission matrix (all APIs)
        </Text>
        <table style={{ borderCollapse: 'collapse', width: '100%' }}>
          <thead>
            <tr style={{ background: '#f3f2f1', textAlign: 'left' }}>
              <th style={cellStyle}>API</th>
              <th style={cellStyle}>Delegated scope</th>
              <th style={cellStyle}>App scope</th>
              <th style={cellStyle}>Admin consent</th>
            </tr>
          </thead>
          <tbody>
            {API_ORDER.map((id) => {
              const api = API_CATALOG[id];
              const isActive = activeApiId !== undefined && id === activeApiId;
              return (
                <tr
                  key={id}
                  style={{
                    background: isActive ? '#deecf9' : '#fff',
                    fontWeight: isActive ? 600 : 400
                  }}
                >
                  <td style={cellStyle}>
                    {isActive && <Icon iconName="ChevronRightSmall" />} {api.title}
                  </td>
                  <td style={cellStyle}>{api.permissions.delegated ?? '—'}</td>
                  <td style={cellStyle}>{api.permissions.application ?? '—'}</td>
                  <td style={cellStyle}>
                    {api.permissions.adminConsent ? (
                      <span style={{ color: '#a4262c' }}>Required</span>
                    ) : (
                      <span style={{ color: '#107c10' }}>No</span>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Stack>
    </div>
  );
};

export default PermissionMatrix;
