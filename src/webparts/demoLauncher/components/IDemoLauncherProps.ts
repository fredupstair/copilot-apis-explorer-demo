import { WebPartContext } from '@microsoft/sp-webpart-base';

/**
 * Props for the root Copilot API Explorer component.
 *
 * We pass the full WebPartContext down because the explorer needs two SPFx
 * services at runtime:
 *   * context.msGraphClientFactory    – delegated MSGraphClientV3 calls.
 *   * context.aadTokenProviderFactory – to obtain the raw delegated access
 *     token purely so we can decode and display its claims in the demo.
 */
export interface IDemoLauncherProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}
