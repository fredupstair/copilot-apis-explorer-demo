// =============================================================================
// RequestBuilder.tsx
// -----------------------------------------------------------------------------
// The "compose the request" section. Two layouts driven by api.method:
//   * POST (Retrieval): an editable JSON body textarea plus a dedicated KQL
//     filter field that the parent injects into the body as `filterExpression`.
//   * GET (everything else): one editable field per query/path/OData parameter.
//
// Both layouts expose "Copy as cURL" and "Copy as Postman Collection" so the
// audience can replay the exact request outside SharePoint.
// =============================================================================

import * as React from 'react';
import {
  DefaultButton,
  Icon,
  MessageBar,
  MessageBarType,
  Stack,
  Text,
  TextField
} from '@fluentui/react';
import { IApiDefinition, IQueryParam } from './types';
import {
  buildCurl,
  buildPostmanCollection,
  copyToClipboard
} from './copilotApi';

export interface IRequestBuilderProps {
  api: IApiDefinition;
  endpointUrl: string;
  /** Token used when generating cURL / Postman exports. */
  activeToken: string;

  // ---- POST (Retrieval) editing ----
  bodyText?: string;
  onBodyTextChange?: (value: string) => void;
  kqlFilter?: string;
  onKqlFilterChange?: (value: string) => void;
  /** Body actually sent on the wire (bodyText merged with the KQL filter). */
  effectiveBody?: unknown;

  // ---- GET editing ----
  params?: IQueryParam[];
  onParamChange?: (index: number, value: string) => void;
}

const RequestBuilder: React.FC<IRequestBuilderProps> = (props) => {
  const {
    api,
    endpointUrl,
    activeToken,
    bodyText,
    onBodyTextChange,
    kqlFilter,
    onKqlFilterChange,
    effectiveBody,
    params,
    onParamChange
  } = props;

  const [copied, setCopied] = React.useState<string | undefined>(undefined);

  const flashCopied = (label: string): void => {
    setCopied(label);
    window.setTimeout(() => setCopied(undefined), 2000);
  };

  const handleCopyCurl = (): void => {
    const curl = buildCurl(api.method, endpointUrl, activeToken, effectiveBody);
    copyToClipboard(curl)
      .then(() => flashCopied('cURL command copied to clipboard'))
      .catch(() => flashCopied('Copy failed – clipboard blocked'));
  };

  const handleCopyPostman = (): void => {
    const collection = buildPostmanCollection(
      api.title,
      api.method,
      endpointUrl,
      activeToken,
      effectiveBody
    );
    copyToClipboard(collection)
      .then(() => flashCopied('Postman v2.1 collection copied to clipboard'))
      .catch(() => flashCopied('Copy failed – clipboard blocked'));
  };

  return (
    <div
      style={{
        border: '1px solid #edebe9',
        borderRadius: 4,
        padding: 16
      }}
    >
      <Stack tokens={{ childrenGap: 12 }}>
        <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
          <Icon iconName="Edit" /> Request builder
        </Text>

        {api.method === 'POST' ? (
          // -------------------- POST: editable JSON + KQL filter --------------------
          <Stack tokens={{ childrenGap: 10 }}>
            <TextField
              label="Request body (JSON)"
              multiline
              rows={10}
              value={bodyText}
              onChange={(_e, v) => onBodyTextChange && onBodyTextChange(v ?? '')}
              styles={{ field: { fontFamily: 'Consolas, monospace', fontSize: 12 } }}
            />
            <TextField
              label='KQL filter — injected into the body as "filterExpression"'
              placeholder={`FileExtension:"docx" OR FileExtension:"pdf"`}
              value={kqlFilter}
              onChange={(_e, v) => onKqlFilterChange && onKqlFilterChange(v ?? '')}
              styles={{ field: { fontFamily: 'Consolas, monospace', fontSize: 12 } }}
            />
          </Stack>
        ) : (
          // -------------------- GET: one field per parameter --------------------
          <Stack tokens={{ childrenGap: 10 }}>
            {(params ?? []).length === 0 && (
              <Text variant="small">This endpoint takes no editable parameters.</Text>
            )}
            {(params ?? []).map((param, index) => (
              <TextField
                key={param.key}
                label={`${param.key}  (${param.location})`}
                description={param.hint}
                value={param.value}
                onChange={(_e, v) => onParamChange && onParamChange(index, v ?? '')}
                styles={{ field: { fontFamily: 'Consolas, monospace', fontSize: 12 } }}
              />
            ))}
          </Stack>
        )}

        <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
          <DefaultButton
            text="Copy as cURL"
            iconProps={{ iconName: 'CommandPrompt' }}
            onClick={handleCopyCurl}
          />
          <DefaultButton
            text="Copy as Postman Collection"
            iconProps={{ iconName: 'Send' }}
            onClick={handleCopyPostman}
          />
        </Stack>

        {copied && (
          <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
            {copied}
          </MessageBar>
        )}
      </Stack>
    </div>
  );
};

export default RequestBuilder;
