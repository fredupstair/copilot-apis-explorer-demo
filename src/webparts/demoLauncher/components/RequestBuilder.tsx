// =============================================================================
// RequestBuilder.tsx
// -----------------------------------------------------------------------------
// The "compose the request" section. Two layouts driven by api.method:
//   * POST (Retrieval): a "Fields to return" checkbox group, an editable JSON
//     body textarea, and a dedicated KQL filter field that the parent injects
//     into the body as `filterExpression`.
//   * GET (everything else): one editable field per query/path/OData parameter.
// =============================================================================

import * as React from 'react';
import { Checkbox, ChoiceGroup, IChoiceGroupOption, Icon, DefaultButton, Stack, Text, TextField } from '@fluentui/react';
import { IApiDefinition, IQueryParam, RetrievalDataSource } from './types';
import { RETRIEVAL_FIELDS } from './copilotApi';

const DATA_SOURCE_OPTIONS: IChoiceGroupOption[] = [
  { key: 'sharePoint', text: 'SharePoint' },
  { key: 'oneDriveBusiness', text: 'OneDrive for Business' },
  { key: 'externalItem', text: 'External item (Copilot connectors)' }
];

/**
 * Curated list of well-known Copilot appClass values. Used to populate the
 * Interaction Export filter checkboxes. The official docs only confirm BizChat
 * and Teams; the others follow the documented IPM.* naming pattern observed in
 * production tenants. Users can edit the URL manually for app classes not
 * listed here.
 */
const APP_CLASS_OPTIONS: { key: string; label: string }[] = [
  { key: 'IPM.SkypeTeams.Message.Copilot.BizChat', label: 'Microsoft 365 Copilot Chat (BizChat)' },
  { key: 'IPM.SkypeTeams.Message.Copilot.Teams', label: 'Copilot in Teams' },
  { key: 'IPM.SkypeTeams.Message.Copilot.Loop', label: 'Copilot in Loop' },
  { key: 'IPM.SkypeTeams.Message.Copilot.OneNote', label: 'Copilot in OneNote' },
  { key: 'IPM.SkypeTeams.Message.Copilot.Whiteboard', label: 'Copilot in Whiteboard' },
  { key: 'IPM.Note.Microsoft.Copilot.Outlook', label: 'Copilot in Outlook' },
  { key: 'IPM.Document.Microsoft.Copilot.Word', label: 'Copilot in Word' },
  { key: 'IPM.Document.Microsoft.Copilot.Excel', label: 'Copilot in Excel' },
  { key: 'IPM.Document.Microsoft.Copilot.PowerPoint', label: 'Copilot in PowerPoint' }
];

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
  /** Dedicated query string input (mirrors body.queryString). */
  queryString?: string;
  onQueryStringChange?: (value: string) => void;
  /** Required dataSource (mirrors body.dataSource). */
  dataSource?: RetrievalDataSource;
  onDataSourceChange?: (value: RetrievalDataSource) => void;
  /** Body actually sent on the wire. */
  effectiveBody?: unknown;
  /** Currently selected metadata fields (resourceMetadata array). */
  selectedFields?: string[];
  onSelectedFieldsChange?: (fields: string[]) => void;

  // ---- GET editing ----
  params?: IQueryParam[];
  onParamChange?: (index: number, value: string) => void;

  // ---- Interaction Export filters (translated into a $filter query param) ----
  interactionAppClasses?: string[];
  onInteractionAppClassesChange?: (values: string[]) => void;
  interactionFrom?: string;
  onInteractionFromChange?: (value: string) => void;
  interactionTo?: string;
  onInteractionToChange?: (value: string) => void;
  /** Auto-built OData $filter expression (preview only). */
  interactionFilter?: string;
}

const RequestBuilder: React.FC<IRequestBuilderProps> = (props) => {
  const {
    api,
    bodyText,
    kqlFilter,
    onKqlFilterChange,
    queryString,
    onQueryStringChange,
    dataSource,
    onDataSourceChange,
    params,
    onParamChange,
    selectedFields,
    onSelectedFieldsChange,
    interactionAppClasses,
    onInteractionAppClassesChange,
    interactionFrom,
    onInteractionFromChange,
    interactionTo,
    onInteractionToChange,
    interactionFilter
  } = props;

  const fields = selectedFields ?? [];
  const allSelected = RETRIEVAL_FIELDS.every((f) => fields.indexOf(f.apiName) !== -1);

  const selectedAppClasses = interactionAppClasses ?? [];
  const toggleAppClass = (key: string, checked: boolean): void => {
    if (!onInteractionAppClassesChange) {
      return;
    }
    if (checked) {
      if (selectedAppClasses.indexOf(key) === -1) {
        onInteractionAppClassesChange([...selectedAppClasses, key]);
      }
    } else {
      onInteractionAppClassesChange(selectedAppClasses.filter((v) => v !== key));
    }
  };

  /** "Last N hours" / "Last N days" shortcut: prefill both from/to using a
   *  rounded-to-the-minute UTC window. */
  const applyRange = (msBack: number): void => {
    if (!onInteractionFromChange || !onInteractionToChange) {
      return;
    }
    const now = new Date();
    const past = new Date(now.getTime() - msBack);
    const fmt = (d: Date): string =>
      `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(
        d.getUTCDate()
      ).padStart(2, '0')}T${String(d.getUTCHours()).padStart(2, '0')}:${String(d.getUTCMinutes()).padStart(2, '0')}`;
    onInteractionFromChange(fmt(past));
    onInteractionToChange(fmt(now));
  };

  const toggleField = (apiName: string, checked: boolean): void => {
    if (!onSelectedFieldsChange) {
      return;
    }
    if (checked) {
      if (fields.indexOf(apiName) === -1) {
        onSelectedFieldsChange([...fields, apiName]);
      }
    } else {
      onSelectedFieldsChange(fields.filter((f) => f !== apiName));
    }
  };

  const toggleAll = (checked: boolean): void => {
    if (!onSelectedFieldsChange) {
      return;
    }
    onSelectedFieldsChange(checked ? RETRIEVAL_FIELDS.map((f) => f.apiName) : []);
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
          // -------------------- POST: structured inputs + readonly JSON preview --------------------
          <Stack tokens={{ childrenGap: 10 }}>
            {api.id === 'retrieval' && (
              <div
                style={{
                  border: '1px solid #edebe9',
                  borderRadius: 4,
                  padding: '8px 10px',
                  background: '#faf9f8'
                }}
              >
                <Stack tokens={{ childrenGap: 4 }}>
                  <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                    Data source <span style={{ color: '#a4262c' }}>*</span>
                  </Text>
                  <Text variant="xSmall" styles={{ root: { color: '#605e5c' } }}>
                    Indicates whether extracts should be retrieved from SharePoint, OneDrive, or Copilot connectors. Acceptable values are <code>sharePoint</code>, <code>oneDriveBusiness</code>, and <code>externalItem</code>. Required.
                  </Text>
                  <ChoiceGroup
                    selectedKey={dataSource}
                    options={DATA_SOURCE_OPTIONS}
                    onChange={(_e, option) =>
                      option && onDataSourceChange && onDataSourceChange(option.key as RetrievalDataSource)
                    }
                    styles={{ flexContainer: { display: 'flex', gap: 16, flexWrap: 'wrap' } }}
                  />
                </Stack>
              </div>
            )}
            <TextField
              label="Query string"
              description='The natural-language query Copilot will ground its answer on (maps to body.queryString).'
              placeholder="Tell me more about Cowork"
              value={queryString ?? ''}
              onChange={(_e, v) => onQueryStringChange && onQueryStringChange(v ?? '')}
              styles={{ field: { fontFamily: 'Consolas, monospace', fontSize: 12 } }}
            />
            <TextField
              label='KQL filter — injected into the body as "filterExpression"'
              placeholder={`FileExtension:"docx" OR FileExtension:"pdf"`}
              value={kqlFilter ?? ''}
              onChange={(_e, v) => onKqlFilterChange && onKqlFilterChange(v ?? '')}
              styles={{ field: { fontFamily: 'Consolas, monospace', fontSize: 12 } }}
            />
            <TextField
              label="Request body (JSON) — read-only preview"
              description="Updated live as you edit the fields above and the checkboxes below."
              multiline
              rows={10}
              readOnly
              value={bodyText}
              styles={{
                field: {
                  fontFamily: 'Consolas, monospace',
                  fontSize: 12,
                  background: '#f3f2f1',
                  color: '#323130'
                }
              }}
            />
            {api.id === 'retrieval' && (
              <div
                style={{
                  border: '1px solid #edebe9',
                  borderRadius: 4,
                  padding: 10,
                  background: '#faf9f8'
                }}
              >
                <Stack tokens={{ childrenGap: 6 }}>
                  <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                    Fields to return
                    <span style={{ fontWeight: 400, color: '#605e5c' }}>
                      {' '}
                      — maps to the body&apos;s <code>resourceMetadata</code> array
                    </span>
                  </Text>
                  <Checkbox
                    label={`Select All (${fields.length}/${RETRIEVAL_FIELDS.length})`}
                    checked={allSelected}
                    indeterminate={!allSelected && fields.length > 0}
                    onChange={(_e, checked) => toggleAll(!!checked)}
                    styles={{ root: { fontWeight: 600 } }}
                  />
                  <div
                    style={{
                      display: 'grid',
                      gridTemplateColumns: 'repeat(auto-fill, minmax(150px, 1fr))',
                      gap: '4px 12px',
                      marginTop: 4
                    }}
                  >
                    {RETRIEVAL_FIELDS.map((field) => (
                      <Checkbox
                        key={field.apiName}
                        label={field.label}
                        checked={fields.indexOf(field.apiName) !== -1}
                        onChange={(_e, checked) => toggleField(field.apiName, !!checked)}
                      />
                    ))}
                  </div>
                </Stack>
              </div>
            )}
          </Stack>
        ) : (
          // -------------------- GET: one field per parameter --------------------
          <Stack tokens={{ childrenGap: 10 }}>
            {(() => {
              // Hide the auto-managed $filter row for Interaction Export so the
              // user only edits it through the structured filter card below.
              const visible = (params ?? [])
                .map((p, i) => ({ p, originalIndex: i }))
                .filter(({ p }) => !(api.id === 'interactionExport' && p.key === '$filter'));
              if (visible.length === 0) {
                return (
                  <Text variant="small">This endpoint takes no editable parameters.</Text>
                );
              }
              return visible.map(({ p, originalIndex }) => (
                <TextField
                  key={p.key}
                  label={`${p.key}  (${p.location})`}
                  description={p.hint}
                  value={p.value}
                  onChange={(_e, v) => onParamChange && onParamChange(originalIndex, v ?? '')}
                  styles={{ field: { fontFamily: 'Consolas, monospace', fontSize: 12 } }}
                />
              ));
            })()}

            {api.id === 'interactionExport' && (
              <div
                style={{
                  border: '1px solid #edebe9',
                  borderRadius: 4,
                  padding: 12,
                  background: '#faf9f8'
                }}
              >
                <Stack tokens={{ childrenGap: 10 }}>
                  <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                    Filters
                    <span style={{ fontWeight: 400, color: '#605e5c' }}>
                      {' '}
                      — translated into the OData{' '}
                      <code>$filter</code> query parameter.
                    </span>
                  </Text>
                  <Text variant="xSmall" styles={{ root: { color: '#605e5c' } }}>
                    Per docs, the <code>createdDateTime</code> filter requires
                    BOTH a minimum and a maximum boundary; date/time fields
                    are treated as UTC.
                  </Text>

                  {/* Created date/time range */}
                  <div
                    style={{
                      display: 'grid',
                      gridTemplateColumns: '1fr 1fr',
                      gap: 12
                    }}
                  >
                    <label style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                      <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                        createdDateTime &gt; (from, UTC)
                      </Text>
                      <input
                        type="datetime-local"
                        value={interactionFrom ?? ''}
                        onChange={(e) =>
                          onInteractionFromChange && onInteractionFromChange(e.target.value)
                        }
                        style={{
                          padding: '6px 8px',
                          border: '1px solid #c8c6c4',
                          borderRadius: 2,
                          fontFamily: 'Consolas, monospace',
                          fontSize: 12,
                          width: '100%'
                        }}
                      />
                    </label>
                    <label style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                      <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                        createdDateTime &lt; (to, UTC)
                      </Text>
                      <input
                        type="datetime-local"
                        value={interactionTo ?? ''}
                        onChange={(e) =>
                          onInteractionToChange && onInteractionToChange(e.target.value)
                        }
                        style={{
                          padding: '6px 8px',
                          border: '1px solid #c8c6c4',
                          borderRadius: 2,
                          fontFamily: 'Consolas, monospace',
                          fontSize: 12,
                          width: '100%'
                        }}
                      />
                    </label>
                  </div>

                  <Stack horizontal tokens={{ childrenGap: 6 }} wrap>
                    <DefaultButton
                      text="Last 24h"
                      onClick={() => applyRange(24 * 60 * 60 * 1000)}
                    />
                    <DefaultButton
                      text="Last 7 days"
                      onClick={() => applyRange(7 * 24 * 60 * 60 * 1000)}
                    />
                    <DefaultButton
                      text="Last 30 days"
                      onClick={() => applyRange(30 * 24 * 60 * 60 * 1000)}
                    />
                    <DefaultButton
                      text="Clear dates"
                      onClick={() => {
                        if (onInteractionFromChange) onInteractionFromChange('');
                        if (onInteractionToChange) onInteractionToChange('');
                      }}
                    />
                  </Stack>

                  {/* App class checkboxes */}
                  <div>
                    <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                      appClass
                      <span style={{ fontWeight: 400, color: '#605e5c' }}>
                        {' '}— select one or more Copilot surfaces
                      </span>
                    </Text>
                    <div
                      style={{
                        display: 'grid',
                        gridTemplateColumns: 'repeat(auto-fill, minmax(220px, 1fr))',
                        gap: '4px 12px',
                        marginTop: 6
                      }}
                    >
                      {APP_CLASS_OPTIONS.map((opt) => (
                        <Checkbox
                          key={opt.key}
                          label={opt.label}
                          checked={selectedAppClasses.indexOf(opt.key) !== -1}
                          onChange={(_e, checked) => toggleAppClass(opt.key, !!checked)}
                          styles={{
                            label: { fontSize: 12 },
                            text: { fontSize: 12 }
                          }}
                        />
                      ))}
                    </div>
                    {selectedAppClasses.length > 0 && (
                      <DefaultButton
                        text="Clear app classes"
                        onClick={() =>
                          onInteractionAppClassesChange && onInteractionAppClassesChange([])
                        }
                        styles={{ root: { marginTop: 8 } }}
                      />
                    )}
                  </div>

                  {/* OData $filter preview */}
                  {interactionFilter && interactionFilter.length > 0 && (
                    <div>
                      <Text variant="xSmall" styles={{ root: { color: '#605e5c' } }}>
                        Auto-generated <code>$filter</code> query parameter:
                      </Text>
                      <pre
                        style={{
                          margin: '4px 0 0 0',
                          padding: 8,
                          background: '#ffffff',
                          border: '1px solid #edebe9',
                          borderRadius: 3,
                          fontFamily: 'Consolas, monospace',
                          fontSize: 11.5,
                          whiteSpace: 'pre-wrap',
                          wordBreak: 'break-word'
                        }}
                      >
                        {interactionFilter}
                      </pre>
                    </div>
                  )}
                </Stack>
              </div>
            )}
          </Stack>
        )}
      </Stack>
    </div>
  );
};

export default RequestBuilder;
