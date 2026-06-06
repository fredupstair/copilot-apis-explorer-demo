// =============================================================================
// ResponseViewer.tsx
// -----------------------------------------------------------------------------
// The "see the result" section. Responsibilities:
//   * Fire the real call (delegates to the parent's onExecute).
//   * Show an HTTP status badge + response time (ms) + size (KB).
//   * Render a *purpose-built* view per API:
//       - Retrieval        -> result cards (title, url, score, excerpt)
//       - Interaction Export -> flat list (timestamp + type + snippet)
//       - Usage APIs       -> pure-CSS bar chart (inline width %)
//   * Always show syntax-highlighted raw JSON underneath, built from React
//     spans (NOT dangerouslySetInnerHTML) so there is no XSS surface.
//   * Copy Response button.
// =============================================================================

import * as React from 'react';
import {
  DefaultButton,
  Icon,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  Stack,
  Text
} from '@fluentui/react';
import {
  ApiId,
  IAiInteraction,
  IApiCallResult,
  IApiDefinition,
  IRetrievalHit
} from './types';
import { copyToClipboard } from './copilotApi';

export interface IResponseViewerProps {
  api: IApiDefinition;
  result: IApiCallResult | undefined;
  executing: boolean;
  onExecute: () => void;
  /** Disable Execute when auth/permissions are not ready, with a reason. */
  canExecute: boolean;
  disabledReason?: string;
}

// -----------------------------------------------------------------------------
// JSON syntax highlighting -> React nodes (no external library, no innerHTML).
// -----------------------------------------------------------------------------

const JSON_TOKEN =
  /("(?:\\.|[^"\\])*"\s*:)|("(?:\\.|[^"\\])*")|\b(true|false|null)\b|(-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)/g;

function highlightJson(value: unknown): React.ReactNode {
  let text: string;
  try {
    text = JSON.stringify(value, undefined, 2);
  } catch {
    text = String(value);
  }
  if (typeof value === 'string') {
    // Non-JSON payloads (e.g. CSV) are shown verbatim.
    text = value;
  }

  const nodes: React.ReactNode[] = [];
  let lastIndex = 0;
  let match: RegExpExecArray | null;
  let key = 0;

  JSON_TOKEN.lastIndex = 0;
  while ((match = JSON_TOKEN.exec(text)) !== null) {
    if (match.index > lastIndex) {
      nodes.push(text.substring(lastIndex, match.index));
    }
    let color = '#001080';
    if (match[1]) {
      color = '#0451a5'; // property key
    } else if (match[2]) {
      color = '#a31515'; // string value
    } else if (match[3]) {
      color = '#0000ff'; // boolean / null
    } else if (match[4]) {
      color = '#098658'; // number
    }
    nodes.push(
      <span key={key++} style={{ color }}>
        {match[0]}
      </span>
    );
    lastIndex = match.index + match[0].length;
  }
  if (lastIndex < text.length) {
    nodes.push(text.substring(lastIndex));
  }
  return nodes;
}

// -----------------------------------------------------------------------------
// Status badge
// -----------------------------------------------------------------------------

const StatusBadge: React.FC<{ status: number; ok: boolean }> = ({ status, ok }) => (
  <span
    style={{
      background: ok ? '#107c10' : status === 0 ? '#605e5c' : '#a4262c',
      color: '#fff',
      borderRadius: 4,
      padding: '2px 10px',
      fontWeight: 700,
      fontSize: 13
    }}
  >
    {status === 0 ? 'NO RESPONSE' : `HTTP ${status}`}
  </span>
);

// -----------------------------------------------------------------------------
// Per-API rich renderers
// -----------------------------------------------------------------------------

function renderRetrieval(body: unknown): React.ReactNode {
  const hits = (body as { retrievalHits?: IRetrievalHit[] })?.retrievalHits;
  if (!hits || hits.length === 0) {
    return <Text variant="small">No retrievalHits returned.</Text>;
  }
  return (
    <Stack tokens={{ childrenGap: 10 }}>
      {hits.map((hit, i) => {
        const title = hit.resourceMetadata?.title || hit.webUrl;
        const topScore = Math.max(0, ...hit.extracts.map((e) => e.relevanceScore ?? 0));
        const excerpt = hit.extracts[0]?.text ?? '';
        return (
          <div
            key={i}
            style={{
              border: '1px solid #edebe9',
              borderLeft: '4px solid #0078d4',
              borderRadius: 4,
              padding: 12,
              background: '#fff'
            }}
          >
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
              <Text styles={{ root: { fontWeight: 600 } }}>{title}</Text>
              {topScore > 0 && (
                <span
                  style={{
                    background: '#deecf9',
                    color: '#004578',
                    borderRadius: 12,
                    padding: '2px 10px',
                    fontSize: 12,
                    fontWeight: 600
                  }}
                >
                  score {topScore.toFixed(3)}
                </span>
              )}
            </Stack>
            <a href={hit.webUrl} target="_blank" rel="noreferrer" style={{ fontSize: 12 }}>
              {hit.webUrl}
            </a>
            <Text variant="small" styles={{ root: { marginTop: 6, color: '#605e5c' } }}>
              {excerpt}
            </Text>
            {hit.sensitivityLabel?.displayName && (
              <span
                style={{
                  display: 'inline-block',
                  marginTop: 6,
                  fontSize: 11,
                  color: hit.sensitivityLabel.color ?? '#605e5c',
                  fontWeight: 600
                }}
              >
                <Icon iconName="Lock" /> {hit.sensitivityLabel.displayName}
              </span>
            )}
          </div>
        );
      })}
    </Stack>
  );
}

function renderInteractions(body: unknown): React.ReactNode {
  const items = (body as { value?: IAiInteraction[] })?.value;
  if (!items || items.length === 0) {
    return <Text variant="small">No interactions returned.</Text>;
  }
  return (
    <Stack tokens={{ childrenGap: 6 }}>
      {items.map((item) => {
        const content = item.body?.content ?? '';
        const snippet = content.length > 160 ? `${content.substring(0, 160)}…` : content;
        const isPrompt = item.interactionType === 'userPrompt';
        return (
          <div
            key={item.id}
            style={{
              border: '1px solid #edebe9',
              borderRadius: 4,
              padding: '8px 12px',
              background: '#fff',
              display: 'flex',
              gap: 12
            }}
          >
            <span
              style={{
                background: isPrompt ? '#fff4ce' : '#deecf9',
                color: isPrompt ? '#8a6d00' : '#004578',
                borderRadius: 4,
                padding: '2px 8px',
                fontSize: 11,
                fontWeight: 600,
                whiteSpace: 'nowrap',
                height: 'fit-content'
              }}
            >
              {item.interactionType ?? 'interaction'}
            </span>
            <div style={{ flex: 1 }}>
              <Text variant="xSmall" styles={{ root: { color: '#605e5c' } }}>
                {item.createdDateTime} · {item.appClass}
              </Text>
              <div style={{ fontSize: 13 }}>{snippet || '(no text content)'}</div>
            </div>
          </div>
        );
      })}
    </Stack>
  );
}

/** Extract {label, value} pairs from an arbitrary usage-report payload. */
function toChartData(body: unknown): { label: string; value: number }[] {
  const data: { label: string; value: number }[] = [];
  const labelKeys = ['reportDate', 'reportPeriod', 'userPrincipalName', 'displayName', 'date'];

  const pickLabel = (row: Record<string, unknown>, fallback: string): string => {
    for (const k of labelKeys) {
      if (typeof row[k] === 'string') {
        return row[k] as string;
      }
    }
    return fallback;
  };

  const pickNumber = (row: Record<string, unknown>): number | undefined => {
    for (const [k, v] of Object.entries(row)) {
      if (labelKeys.indexOf(k) !== -1) {
        continue;
      }
      const n = typeof v === 'number' ? v : typeof v === 'string' ? Number(v) : NaN;
      if (!isNaN(n) && v !== '' && v !== undefined) {
        return n;
      }
    }
    return undefined;
  };

  const value = (body as { value?: Record<string, unknown>[] })?.value;
  if (Array.isArray(value)) {
    value.forEach((row, i) => {
      const n = pickNumber(row);
      if (n !== undefined) {
        data.push({ label: pickLabel(row, `#${i + 1}`), value: n });
      }
    });
  } else if (body && typeof body === 'object') {
    // Single summary object: each numeric property becomes a bar.
    for (const [k, v] of Object.entries(body as Record<string, unknown>)) {
      const n = typeof v === 'number' ? v : typeof v === 'string' ? Number(v) : NaN;
      if (!isNaN(n) && v !== '' && k.indexOf('@odata') === -1) {
        data.push({ label: k, value: n });
      }
    }
  }
  return data;
}

function renderBarChart(body: unknown): React.ReactNode {
  const data = toChartData(body);
  if (data.length === 0) {
    return (
      <Text variant="small">
        No numeric series detected to chart. Inspect the raw JSON below.
      </Text>
    );
  }
  const max = Math.max(...data.map((d) => d.value), 1);
  return (
    <Stack tokens={{ childrenGap: 6 }}>
      {data.slice(0, 30).map((d, i) => (
        <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <span
            style={{
              width: 150,
              fontSize: 12,
              overflow: 'hidden',
              textOverflow: 'ellipsis',
              whiteSpace: 'nowrap'
            }}
            title={d.label}
          >
            {d.label}
          </span>
          <div style={{ flex: 1, background: '#f3f2f1', borderRadius: 4, height: 18 }}>
            <div
              style={{
                // Pure-CSS bar: width is the value as a percentage of the max.
                width: `${Math.max(2, (d.value / max) * 100)}%`,
                background: '#0078d4',
                height: '100%',
                borderRadius: 4
              }}
            />
          </div>
          <span style={{ width: 60, textAlign: 'right', fontSize: 12, fontWeight: 600 }}>
            {d.value}
          </span>
        </div>
      ))}
    </Stack>
  );
}

const RICH_RENDERERS: Record<ApiId, (body: unknown) => React.ReactNode> = {
  retrieval: renderRetrieval,
  interactionExport: renderInteractions,
  usageSummary: renderBarChart,
  usageTrend: renderBarChart,
  usageUserDetail: renderBarChart
};

// -----------------------------------------------------------------------------
// Component
// -----------------------------------------------------------------------------

const ResponseViewer: React.FC<IResponseViewerProps> = (props) => {
  const { api, result, executing, onExecute, canExecute, disabledReason } = props;
  const [copied, setCopied] = React.useState<boolean>(false);

  const handleCopy = (): void => {
    if (!result) {
      return;
    }
    const text =
      typeof result.body === 'string'
        ? result.body
        : JSON.stringify(result.body, undefined, 2);
    copyToClipboard(text)
      .then(() => {
        setCopied(true);
        window.setTimeout(() => setCopied(false), 2000);
      })
      .catch(() => undefined);
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
        <Stack horizontal verticalAlign="center" horizontalAlign="space-between">
          <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
            <Icon iconName="Play" /> Response
          </Text>
          <PrimaryButton
            text={executing ? 'Executing…' : 'Execute'}
            iconProps={{ iconName: 'Play' }}
            disabled={executing || !canExecute}
            onClick={onExecute}
          />
        </Stack>

        {!canExecute && disabledReason && (
          <MessageBar messageBarType={MessageBarType.warning}>{disabledReason}</MessageBar>
        )}

        {executing && <Spinner size={SpinnerSize.medium} label="Calling Microsoft Graph…" />}

        {result && !executing && (
          <Stack tokens={{ childrenGap: 12 }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }} wrap>
              <StatusBadge status={result.status} ok={result.ok} />
              <Text variant="small">
                <Icon iconName="Timer" /> {result.timeMs} ms
              </Text>
              <Text variant="small">
                <Icon iconName="Download" /> {result.sizeKb} KB
              </Text>
              <DefaultButton
                text={copied ? 'Copied!' : 'Copy Response'}
                iconProps={{ iconName: 'Copy' }}
                onClick={handleCopy}
              />
            </Stack>

            {result.error && (
              <MessageBar messageBarType={MessageBarType.error}>{result.error}</MessageBar>
            )}

            {/* Purpose-built rendering for successful calls. */}
            {result.ok && (
              <div>
                <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, marginBottom: 6 } }}>
                  Rendered view
                </Text>
                {RICH_RENDERERS[api.id](result.body)}
              </div>
            )}

            {/* Raw, syntax-highlighted JSON. */}
            <div>
              <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, marginBottom: 6 } }}>
                Raw response
              </Text>
              <pre
                style={{
                  background: '#1e1e1e00',
                  border: '1px solid #edebe9',
                  borderRadius: 4,
                  padding: 12,
                  margin: 0,
                  maxHeight: 360,
                  overflow: 'auto',
                  fontFamily: 'Consolas, monospace',
                  fontSize: 12,
                  lineHeight: 1.5,
                  whiteSpace: 'pre-wrap',
                  wordBreak: 'break-word'
                }}
              >
                {highlightJson(result.body)}
              </pre>
            </div>
          </Stack>
        )}
      </Stack>
    </div>
  );
};

export default ResponseViewer;
