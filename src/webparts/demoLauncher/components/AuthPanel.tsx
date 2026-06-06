// =============================================================================
// AuthPanel.tsx
// -----------------------------------------------------------------------------
// Sidebar authentication surface. Two modes:
//   1. Delegated  (MSGraphClientV3) – the signed-in SharePoint user. The raw
//      access token is also acquired (on demand) purely so we can decode and
//      display its claims for the audience.
//   2. App Registration (Application) – an app-only token obtained OUTSIDE the
//      browser (e.g. via the auth/Get-AppToken.ps1 helper) and pasted into the
//      textbox below. The browser CANNOT safely run the client_credentials
//      flow itself: it is blocked by CORS and would expose the client secret.
//
// Whatever the mode, we decode the active JWT client-side (no library) and show
// upn, scopes/roles and a live expiry countdown, plus a green/red permission
// check for the currently selected API.
// =============================================================================

import * as React from 'react';
import {
  ChoiceGroup,
  DefaultButton,
  IChoiceGroupOption,
  Icon,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Stack,
  Text,
  TextField
} from '@fluentui/react';
import { AuthMode, IDecodedToken, IPermissionCheck } from './types';
import { formatExpiryCountdown } from './copilotApi';

export interface IAuthPanelProps {
  authMode: AuthMode;
  onAuthModeChange: (mode: AuthMode) => void;

  /** Delegated token (acquired on demand) – used only for display/decoding. */
  delegatedToken: string;
  onAcquireDelegatedToken: () => void;
  acquiringDelegated: boolean;

  /** App-only token pasted by the user (obtained out-of-band). */
  appToken: string;
  onAppTokenChange: (token: string) => void;

  /** Decoded claims for whichever token is currently active. */
  decoded: IDecodedToken | undefined;

  /** Permission check for the active API + auth mode. */
  permissionCheck: IPermissionCheck;
  selectedApiTitle: string;
}

const modeOptions: IChoiceGroupOption[] = [
  { key: 'graphClient', text: 'Delegated (current user)' },
  { key: 'manualToken', text: 'App Registration (Application)' }
];

const sectionStyle: React.CSSProperties = {
  background: '#ffffff',
  padding: 12,
  borderRadius: 4,
  border: '1px solid #edebe9'
};

const AuthPanel: React.FC<IAuthPanelProps> = (props) => {
  const {
    authMode,
    onAuthModeChange,
    delegatedToken,
    onAcquireDelegatedToken,
    acquiringDelegated,
    appToken,
    onAppTokenChange,
    decoded,
    permissionCheck,
    selectedApiTitle
  } = props;

  // A 1-second ticker so the expiry countdown actually counts down on screen.
  const [, setTick] = React.useState<number>(0);
  React.useEffect(() => {
    const handle = window.setInterval(() => setTick((t) => t + 1), 1000);
    return () => window.clearInterval(handle);
  }, []);

  const activeToken = authMode === 'manualToken' ? appToken : delegatedToken;

  return (
    <div
      style={{
        border: '1px solid #edebe9',
        borderRadius: 4,
        padding: 16,
        background: '#faf9f8',
        height: '100%',
        boxSizing: 'border-box'
      }}
    >
      <Stack tokens={{ childrenGap: 12 }}>
        <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
          <Icon iconName="Permissions" /> Authentication
        </Text>

        <ChoiceGroup
          selectedKey={authMode}
          options={modeOptions}
          onChange={(_e, option) => option && onAuthModeChange(option.key as AuthMode)}
        />

        {authMode === 'graphClient' ? (
          // ------------------------- Delegated mode -------------------------
          <Stack tokens={{ childrenGap: 8 }}>
            <Text variant="small">
              Calls run as the signed-in user via{' '}
              <code>msGraphClientFactory.getClient(&apos;3&apos;)</code>. Use the
              button below to also acquire the raw access token and inspect its
              claims.
            </Text>
            <PrimaryButton
              text={acquiringDelegated ? 'Acquiring…' : 'Acquire & decode delegated token'}
              iconProps={{ iconName: 'Signin' }}
              disabled={acquiringDelegated}
              onClick={onAcquireDelegatedToken}
            />
          </Stack>
        ) : (
          // ------------------------- App Registration mode -------------------------
          <Stack tokens={{ childrenGap: 8 }}>
            <MessageBar messageBarType={MessageBarType.info} isMultiline>
              The browser cannot run the <code>client_credentials</code> flow
              (CORS blocks the AAD token endpoint, and a client secret must
              never live in JavaScript). Generate the token out-of-band — for
              example with <code>auth/Get-AppToken.ps1</code> — and paste it
              below. The countdown shows the JWT&apos;s <code>exp</code> claim.
            </MessageBar>

            <TextField
              label="App-only access token"
              multiline
              rows={4}
              resizable={false}
              autoAdjustHeight={false}
              placeholder="Paste an app-only JWT (eyJ0...)"
              value={appToken}
              onChange={(_e, newValue) => onAppTokenChange(newValue ?? '')}
              styles={{
                field: {
                  fontFamily: 'Consolas, monospace',
                  fontSize: 12,
                  wordBreak: 'break-all'
                }
              }}
            />

            {appToken && (
              <DefaultButton
                text="Clear token"
                iconProps={{ iconName: 'Clear' }}
                onClick={() => onAppTokenChange('')}
              />
            )}
          </Stack>
        )}

        {/* Decoded JWT summary */}
        {activeToken && !decoded && (
          <MessageBar messageBarType={MessageBarType.error}>
            The provided value is not a valid three-part JWT and could not be decoded.
          </MessageBar>
        )}

        {decoded && (
          <div style={sectionStyle}>
            <Stack tokens={{ childrenGap: 6 }}>
              <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
                Decoded token{' '}
                <span style={{ fontWeight: 400, color: decoded.isAppToken ? '#8764b8' : '#0078d4' }}>
                  ({decoded.isAppToken ? 'Application' : 'Delegated'})
                </span>
              </Text>
              <Text variant="small">
                <strong>upn / identity:</strong> {decoded.upn ?? (decoded.isAppToken ? '(app token – no user)' : 'n/a')}
              </Text>
              <Text variant="small">
                <strong>tenant:</strong> {decoded.tenantId ?? 'n/a'}
              </Text>
              <Text variant="small">
                <strong>audience:</strong> {decoded.audience ?? 'n/a'}
              </Text>
              <Text variant="small">
                <strong>expires in:</strong>{' '}
                <span
                  style={{
                    fontWeight: 600,
                    color: formatExpiryCountdown(decoded.exp) === 'EXPIRED' ? '#a4262c' : '#107c10'
                  }}
                >
                  {formatExpiryCountdown(decoded.exp)}
                </span>
              </Text>
              <Text variant="small">
                <strong>{decoded.isAppToken ? 'roles' : 'scopes'}:</strong>{' '}
                {(decoded.isAppToken ? decoded.roles : decoded.scopes).join(', ') || '(none)'}
              </Text>
            </Stack>
          </div>
        )}

        {/* Permission check for the selected API */}
        <div style={sectionStyle}>
          <Stack tokens={{ childrenGap: 4 }}>
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
              Permission check · {selectedApiTitle}
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
      </Stack>
    </div>
  );
};

export default AuthPanel;
