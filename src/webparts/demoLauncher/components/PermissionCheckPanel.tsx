// =============================================================================
// PermissionCheckPanel.tsx
// -----------------------------------------------------------------------------
// Per-tab "Permission check · <API title>" card. Shown above the request
// builder on every API tab so the audience can immediately see whether the
// active token carries the scopes/roles required by the selected endpoint.
// =============================================================================

import * as React from 'react';
import { Icon, MessageBar, MessageBarType, Stack, Text } from '@fluentui/react';
import { IPermissionCheck } from './types';

export interface IPermissionCheckPanelProps {
  apiTitle: string;
  permissionCheck: IPermissionCheck;
}

const PermissionCheckPanel: React.FC<IPermissionCheckPanelProps> = ({
  apiTitle,
  permissionCheck
}) => (
  <div
    style={{
      border: '1px solid #edebe9',
      borderRadius: 4,
      padding: 12,
      background: '#ffffff'
    }}
  >
    <Stack tokens={{ childrenGap: 6 }}>
      <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
        <Icon iconName="Permissions" /> Permission check · {apiTitle}
      </Text>
      {permissionCheck.required.length === 0 ? (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline>
          This API is not supported for the current auth mode (see the info panel).
        </MessageBar>
      ) : (
        <Stack horizontal wrap tokens={{ childrenGap: 6 }}>
          {permissionCheck.required.map((scope) => {
            const ok = permissionCheck.granted.indexOf(scope) !== -1;
            return (
              <span
                key={scope}
                style={{
                  padding: '2px 10px',
                  borderRadius: 12,
                  fontSize: 12,
                  fontWeight: 600,
                  color: '#fff',
                  background: ok ? '#107c10' : '#a4262c'
                }}
              >
                <Icon iconName={ok ? 'CheckMark' : 'Cancel'} /> {scope}
              </span>
            );
          })}
        </Stack>
      )}
    </Stack>
  </div>
);

export default PermissionCheckPanel;
