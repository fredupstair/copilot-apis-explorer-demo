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
  Modal,
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
  IRetrievalExtract,
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
// JSON helpers for the Raw response viewer.
// Everything is rendered as React nodes (no external library, no innerHTML).
// -----------------------------------------------------------------------------

function stringifySafe(value: unknown): string {
  if (typeof value === 'string') {
    return value;
  }
  try {
    return JSON.stringify(value, undefined, 2);
  } catch {
    return String(value);
  }
}

// -----------------------------------------------------------------------------
// Collapsible JSON tree builder for the Raw response viewer.
// Walks the value into a flat list of rows. Each object/array becomes a row
// for its opening brace (with a [+]/[-] toggle), followed by its children, then
// a closing-brace row. When a path is collapsed we emit a single condensed
// row with a "… N keys" placeholder.
// -----------------------------------------------------------------------------

interface IJsonRow {
  indent: number;
  toggle?: { collapsed: boolean; onToggle: () => void; path: string };
  content: React.ReactNode;
}

function renderJsonPrimitive(value: unknown, k: number): React.ReactNode {
  if (typeof value === 'string') {
    return (
      <span key={k} style={{ color: '#a31515' }}>
        {JSON.stringify(value)}
      </span>
    );
  }
  if (typeof value === 'number') {
    return (
      <span key={k} style={{ color: '#098658' }}>
        {String(value)}
      </span>
    );
  }
  if (typeof value === 'boolean' || value === null) {
    return (
      <span key={k} style={{ color: '#0000ff' }}>
        {String(value)}
      </span>
    );
  }
  // Fallback for things like undefined or symbols (shouldn't happen in JSON).
  return (
    <span key={k} style={{ color: '#605e5c' }}>
      {String(value)}
    </span>
  );
}

function renderKeyLabel(key: string, k: number): React.ReactNode {
  return (
    <React.Fragment key={k}>
      <span style={{ color: '#0451a5' }}>{JSON.stringify(key)}</span>
      <span>: </span>
    </React.Fragment>
  );
}

function collectCollapsiblePaths(value: unknown, path: string, out: string[]): void {
  if (value === null || typeof value !== 'object') {
    return;
  }
  if (Array.isArray(value)) {
    if (value.length > 0) {
      out.push(path);
      value.forEach((v, i) => collectCollapsiblePaths(v, path + '/' + i, out));
    }
    return;
  }
  const entries = Object.entries(value as Record<string, unknown>);
  if (entries.length > 0) {
    out.push(path);
    entries.forEach(([k, v]) => collectCollapsiblePaths(v, path + '/' + k, out));
  }
}

function buildJsonRows(
  value: unknown,
  collapsedPaths: Set<string>,
  toggle: (path: string) => void
): IJsonRow[] {
  const rows: IJsonRow[] = [];
  let nodeKey = 0;

  function emit(
    val: unknown,
    path: string,
    indent: number,
    keyLabel: React.ReactNode | null,
    trailingComma: boolean
  ): void {
    const isObject = val !== null && typeof val === 'object' && !Array.isArray(val);
    const isArray = Array.isArray(val);

    if (!isObject && !isArray) {
      rows.push({
        indent,
        content: (
          <>
            {keyLabel}
            {renderJsonPrimitive(val, nodeKey++)}
            {trailingComma && <span>,</span>}
          </>
        )
      });
      return;
    }

    const entries = isArray
      ? (val as unknown[]).map((v, i) => [String(i), v] as [string, unknown])
      : Object.entries(val as Record<string, unknown>);
    const open = isArray ? '[' : '{';
    const close = isArray ? ']' : '}';
    const comma = trailingComma ? <span>,</span> : null;

    if (entries.length === 0) {
      rows.push({
        indent,
        content: (
          <>
            {keyLabel}
            <span>{open}</span>
            <span>{close}</span>
            {comma}
          </>
        )
      });
      return;
    }

    const collapsed = collapsedPaths.has(path);
    if (collapsed) {
      const noun = isArray
        ? entries.length === 1
          ? 'item'
          : 'items'
        : entries.length === 1
        ? 'key'
        : 'keys';
      rows.push({
        indent,
        toggle: { collapsed: true, onToggle: () => toggle(path), path },
        content: (
          <>
            {keyLabel}
            <span>{open}</span>
            <span style={{ color: '#a8a8a8', fontStyle: 'italic', margin: '0 6px' }}>
              … {entries.length} {noun}
            </span>
            <span>{close}</span>
            {comma}
          </>
        )
      });
      return;
    }

    rows.push({
      indent,
      toggle: { collapsed: false, onToggle: () => toggle(path), path },
      content: (
        <>
          {keyLabel}
          <span>{open}</span>
        </>
      )
    });
    entries.forEach(([k, v], idx) => {
      const childLabel: React.ReactNode | null = isArray ? null : renderKeyLabel(k, nodeKey++);
      const childPath = path + '/' + k;
      emit(v, childPath, indent + 1, childLabel, idx < entries.length - 1);
    });
    rows.push({
      indent,
      content: (
        <>
          <span>{close}</span>
          {comma}
        </>
      )
    });
  }

  emit(value, '', 0, null, false);
  return rows;
}

// -----------------------------------------------------------------------------
// RawJsonViewer
// Pretty-printed, syntax-highlighted JSON with a line-number gutter, per-node
// [+]/[-] collapse toggles, and an "Expand" button that opens the same view
// inside a near-fullscreen Fluent UI Modal. The inline view caps the height;
// the modal view grows to fill the viewport. Toolbar offers Collapse all /
// Expand all helpers and a Copy button (always copies the *full* JSON).
// -----------------------------------------------------------------------------

const INDENT_PX = 14;
const TOGGLE_SLOT_PX = 16;

const RawJsonViewer: React.FC<{ value: unknown }> = ({ value }) => {
  const [expanded, setExpanded] = React.useState<boolean>(false);
  const [copied, setCopied] = React.useState<boolean>(false);
  const [collapsedPaths, setCollapsedPaths] = React.useState<Set<string>>(new Set());

  // When the underlying value changes, drop any stale collapse state.
  React.useEffect(() => {
    setCollapsedPaths(new Set());
  }, [value]);

  const togglePath = React.useCallback((path: string) => {
    setCollapsedPaths((prev) => {
      const next = new Set(prev);
      if (next.has(path)) {
        next.delete(path);
      } else {
        next.add(path);
      }
      return next;
    });
  }, []);

  const rows = React.useMemo(
    () => buildJsonRows(value, collapsedPaths, togglePath),
    [value, collapsedPaths, togglePath]
  );

  const plainText = React.useMemo(() => stringifySafe(value), [value]);
  const sizeKb = React.useMemo(() => (new Blob([plainText]).size / 1024).toFixed(1), [plainText]);

  const handleCopy = (): void => {
    copyToClipboard(plainText)
      .then(() => {
        setCopied(true);
        window.setTimeout(() => setCopied(false), 1500);
      })
      .catch(() => undefined);
  };

  const handleCollapseAll = (): void => {
    const all: string[] = [];
    collectCollapsiblePaths(value, '', all);
    // Don't collapse the root – otherwise the whole tree disappears behind
    // a single "{ … }" row, which is rarely what the user wants.
    setCollapsedPaths(new Set(all.filter((p) => p !== '')));
  };

  const handleExpandAll = (): void => {
    setCollapsedPaths(new Set());
  };

  const gutterWidth = Math.max(34, String(rows.length).length * 9 + 20);

  const codeBlock = (maxHeight: number | string): React.ReactNode => (
    <div
      style={{
        border: '1px solid #edebe9',
        borderRadius: 4,
        background: '#fbfbfb',
        overflow: 'auto',
        maxHeight,
        height: maxHeight === '100%' ? '100%' : undefined,
        fontFamily: 'Consolas, "Courier New", monospace',
        fontSize: 12.5,
        lineHeight: '1.6em'
      }}
    >
      <table
        style={{
          borderCollapse: 'collapse',
          width: '100%',
          tableLayout: 'fixed'
        }}
      >
        <colgroup>
          <col style={{ width: gutterWidth }} />
          <col />
        </colgroup>
        <tbody>
          {rows.map((row, idx) => (
            <tr key={idx}>
              <td
                style={{
                  textAlign: 'right',
                  padding: '0 8px 0 6px',
                  color: '#a8a8a8',
                  background: '#f3f2f1',
                  borderRight: '1px solid #edebe9',
                  userSelect: 'none',
                  verticalAlign: 'top',
                  fontVariantNumeric: 'tabular-nums',
                  position: 'sticky',
                  left: 0
                }}
              >
                {idx + 1}
              </td>
              <td
                style={{
                  padding: '0 12px',
                  whiteSpace: 'pre',
                  verticalAlign: 'top'
                }}
              >
                <span
                  style={{
                    display: 'inline-block',
                    width: row.indent * INDENT_PX
                  }}
                />
                <span
                  style={{
                    display: 'inline-block',
                    width: TOGGLE_SLOT_PX,
                    textAlign: 'center'
                  }}
                >
                  {row.toggle ? (
                    <button
                      type="button"
                      onClick={row.toggle.onToggle}
                      aria-expanded={!row.toggle.collapsed}
                      aria-label={row.toggle.collapsed ? 'Expand node' : 'Collapse node'}
                      title={row.toggle.collapsed ? 'Expand' : 'Collapse'}
                      style={{
                        width: 14,
                        height: 14,
                        lineHeight: '12px',
                        padding: 0,
                        margin: 0,
                        border: '1px solid #c8c6c4',
                        background: '#ffffff',
                        color: '#323130',
                        fontFamily: 'inherit',
                        fontSize: 11,
                        fontWeight: 700,
                        cursor: 'pointer',
                        borderRadius: 2,
                        display: 'inline-flex',
                        alignItems: 'center',
                        justifyContent: 'center'
                      }}
                    >
                      {row.toggle.collapsed ? '+' : '\u2212'}
                    </button>
                  ) : null}
                </span>
                {row.content}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );

  const toolbar = (inModal: boolean): React.ReactNode => (
    <Stack
      horizontal
      verticalAlign="center"
      tokens={{ childrenGap: 8 }}
      wrap
      styles={{ root: { marginBottom: inModal ? 0 : 6 } }}
    >
      <Text
        variant={inModal ? 'large' : 'mediumPlus'}
        styles={{ root: { fontWeight: 600, flexGrow: 1 } }}
      >
        <Icon iconName="Code" /> Raw response
      </Text>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        {rows.length} {rows.length === 1 ? 'line' : 'lines'} · {sizeKb} KB
      </Text>
      <DefaultButton
        text="Collapse all"
        iconProps={{ iconName: 'CollapseContent' }}
        onClick={handleCollapseAll}
      />
      <DefaultButton
        text="Expand all"
        iconProps={{ iconName: 'ExploreContent' }}
        onClick={handleExpandAll}
      />
      <DefaultButton
        text={copied ? 'Copied!' : 'Copy'}
        iconProps={{ iconName: 'Copy' }}
        onClick={handleCopy}
      />
      {inModal ? (
        <DefaultButton
          text="Close"
          iconProps={{ iconName: 'ChromeClose' }}
          onClick={() => setExpanded(false)}
        />
      ) : (
        <DefaultButton
          text="Expand"
          iconProps={{ iconName: 'FullScreen' }}
          onClick={() => setExpanded(true)}
        />
      )}
    </Stack>
  );

  return (
    <div>
      {toolbar(false)}
      {codeBlock(420)}
      <Modal
        isOpen={expanded}
        onDismiss={() => setExpanded(false)}
        isBlocking={false}
        styles={{
          main: {
            width: '95vw',
            maxWidth: '95vw',
            height: '92vh',
            maxHeight: '92vh',
            display: 'flex',
            flexDirection: 'column'
          },
          scrollableContent: {
            display: 'flex',
            flexDirection: 'column',
            flexGrow: 1,
            overflow: 'hidden'
          }
        }}
      >
        <Stack
          tokens={{ childrenGap: 12 }}
          styles={{
            root: { padding: 16, height: '100%', boxSizing: 'border-box' }
          }}
        >
          {toolbar(true)}
          <div style={{ flexGrow: 1, minHeight: 0, display: 'flex' }}>
            <div style={{ flexGrow: 1, minHeight: 0, display: 'flex' }}>
              {codeBlock('100%')}
            </div>
          </div>
        </Stack>
      </Modal>
    </div>
  );
};

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
  // Rank hits by best extract score, descending – mirrors how Copilot picks
  // grounding evidence and keeps the most relevant doc on top.
  const ranked = hits
    .map((hit, originalIdx) => ({
      hit,
      originalIdx,
      topScore: Math.max(0, ...hit.extracts.map((e) => e.relevanceScore ?? 0))
    }))
    .sort((a, b) => b.topScore - a.topScore);

  const totalExtracts = hits.reduce((sum, h) => sum + h.extracts.length, 0);

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center" wrap>
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
          <Icon iconName="Search" /> {hits.length} {hits.length === 1 ? 'hit' : 'hits'} · {totalExtracts}{' '}
          {totalExtracts === 1 ? 'extract' : 'extracts'}
        </span>
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          Sorted by best extract relevance score.
        </Text>
      </Stack>
      {ranked.map(({ hit, topScore, originalIdx }, rankIdx) => (
        <RetrievalHitCard
          key={originalIdx}
          hit={hit}
          rank={rankIdx + 1}
          topScore={topScore}
        />
      ))}
    </Stack>
  );
}

/** A coloured score chip: green ≥0.7, amber ≥0.5, red below. */
const ScoreChip: React.FC<{ score: number; label?: string }> = ({ score, label }) => {
  let bg = '#fde7e9';
  let color = '#a4262c';
  if (score >= 0.7) {
    bg = '#dff6dd';
    color = '#0b6a0b';
  } else if (score >= 0.5) {
    bg = '#fff4ce';
    color = '#8a6d00';
  }
  return (
    <span
      style={{
        background: bg,
        color,
        borderRadius: 12,
        padding: '2px 10px',
        fontSize: 12,
        fontWeight: 600,
        whiteSpace: 'nowrap'
      }}
      title={`Relevance score: ${score.toFixed(4)}`}
    >
      {label ?? `score ${score.toFixed(3)}`}
    </span>
  );
};

/** Friendly icon per resourceType returned by the Retrieval API. */
function resourceTypeIcon(resourceType?: string): string {
  switch ((resourceType ?? '').toLowerCase()) {
    case 'listitem':
      return 'Page';
    case 'driveitem':
      return 'Document';
    case 'message':
      return 'Mail';
    case 'chatmessage':
      return 'Chat';
    case 'event':
      return 'Calendar';
    default:
      return 'FileAsk';
  }
}

/** Strip the leading host + collapse "/sites/<x>/..." into a short crumb. */
function describeUrl(url: string): { siteCrumb: string; pageName: string } {
  try {
    const parsed = new URL(url);
    const segments = parsed.pathname.split('/').filter((s) => s.length > 0);
    const pageName = decodeURIComponent(segments[segments.length - 1] ?? parsed.pathname).replace(
      /\.aspx$/i,
      ''
    );
    const sitesIdx = segments.indexOf('sites');
    let siteCrumb = parsed.hostname;
    if (sitesIdx !== -1 && segments[sitesIdx + 1]) {
      siteCrumb = `${parsed.hostname} · ${segments[sitesIdx + 1]}`;
    }
    return { siteCrumb, pageName };
  } catch {
    return { siteCrumb: url, pageName: url };
  }
}

/** A single retrieval hit card with collapsible extracts. */
function RetrievalHitCard(props: { hit: IRetrievalHit; rank: number; topScore: number }): JSX.Element {
  const { hit, rank, topScore } = props;
  const title = hit.resourceMetadata?.title || hit.webUrl;
  const author = hit.resourceMetadata?.author;
  const { siteCrumb, pageName } = describeUrl(hit.webUrl);

  return (
    <div
      style={{
        border: '1px solid #edebe9',
        borderLeft: '4px solid #0078d4',
        borderRadius: 4,
        padding: 14,
        background: '#fff'
      }}
    >
      <Stack tokens={{ childrenGap: 8 }}>
        {/* Header: rank + title + score */}
        <Stack horizontal horizontalAlign="space-between" verticalAlign="start" tokens={{ childrenGap: 12 }}>
          <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center" styles={{ root: { minWidth: 0 } }}>
            <span
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                justifyContent: 'center',
                minWidth: 24,
                height: 24,
                borderRadius: 12,
                background: '#0078d4',
                color: '#fff',
                fontSize: 12,
                fontWeight: 700,
                padding: '0 8px'
              }}
            >
              #{rank}
            </span>
            <Icon iconName={resourceTypeIcon(hit.resourceType)} style={{ color: '#0078d4' }} />
            <Text
              styles={{
                root: {
                  fontWeight: 600,
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                  whiteSpace: 'nowrap'
                }
              }}
              title={title}
            >
              {title}
            </Text>
          </Stack>
          <ScoreChip score={topScore} />
        </Stack>

        {/* Meta: site/page + author + resource type + extract count */}
        <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
          <span style={{ fontSize: 11, color: '#605e5c' }}>
            <Icon iconName="SharepointAppIcon16" /> {siteCrumb} / <strong>{pageName}</strong>
          </span>
          {author && (
            <span style={{ fontSize: 11, color: '#605e5c' }}>
              <Icon iconName="Contact" /> {author}
            </span>
          )}
          {hit.resourceType && (
            <span
              style={{
                fontSize: 11,
                color: '#323130',
                background: '#f3f2f1',
                padding: '0 6px',
                borderRadius: 4,
                fontWeight: 600
              }}
            >
              {hit.resourceType}
            </span>
          )}
          <span style={{ fontSize: 11, color: '#605e5c' }}>
            {hit.extracts.length} {hit.extracts.length === 1 ? 'extract' : 'extracts'}
          </span>
        </Stack>

        <a
          href={hit.webUrl}
          target="_blank"
          rel="noreferrer"
          style={{ fontSize: 11, wordBreak: 'break-all' }}
        >
          {hit.webUrl}
        </a>

        {hit.sensitivityLabel?.displayName && (
          <span
            style={{
              display: 'inline-block',
              fontSize: 11,
              color: hit.sensitivityLabel.color ?? '#605e5c',
              fontWeight: 600
            }}
          >
            <Icon iconName="Lock" /> {hit.sensitivityLabel.displayName}
          </span>
        )}

        {/* Extracts list */}
        <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 4 } }}>
          {hit.extracts.map((extract, idx) => (
            <ExtractBlock key={idx} extract={extract} index={idx} total={hit.extracts.length} />
          ))}
        </Stack>
      </Stack>
    </div>
  );
}

/** A single extract rendered as a collapsible accordion. Collapsed by default;
 *  the header stays visible (with the score chip) so the audience can scan
 *  relevance at a glance, and clicking expands the full text. Once expanded
 *  the user can toggle between a Formatted view (markdown + inline styled
 *  spans + slide separators) and a Raw view (verbatim text). */
function ExtractBlock(props: { extract: IRetrievalExtract; index: number; total: number }): JSX.Element {
  const { extract, index, total } = props;
  const [expanded, setExpanded] = React.useState<boolean>(false);
  const [mode, setMode] = React.useState<'formatted' | 'raw'>('formatted');
  const cleaned = cleanExtractText(extract.text);
  const formattedSource = prepareFormattedExtract(cleaned);
  const charCount = cleaned.length;
  const panelId = `extract-${index}-panel`;

  const toggleBtnStyle = (active: boolean): React.CSSProperties => ({
    background: active ? '#0078d4' : '#ffffff',
    color: active ? '#ffffff' : '#323130',
    border: '1px solid ' + (active ? '#0078d4' : '#c8c6c4'),
    padding: '2px 10px',
    fontSize: 11,
    fontWeight: 600,
    cursor: 'pointer',
    lineHeight: '18px'
  });

  return (
    <div
      style={{
        background: '#faf9f8',
        border: '1px solid #edebe9',
        borderRadius: 4,
        overflow: 'hidden'
      }}
    >
      <button
        type="button"
        onClick={() => setExpanded((v) => !v)}
        aria-expanded={expanded}
        aria-controls={panelId}
        style={{
          width: '100%',
          background: 'transparent',
          border: 'none',
          padding: '8px 10px',
          cursor: 'pointer',
          textAlign: 'left',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          gap: 8
        }}
      >
        <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
          <Icon
            iconName="ChevronRight"
            style={{
              fontSize: 10,
              color: '#605e5c',
              transform: expanded ? 'rotate(90deg)' : 'rotate(0deg)',
              transition: 'transform 0.15s ease'
            }}
          />
          <Text variant="xSmall" styles={{ root: { color: '#605e5c', fontWeight: 600 } }}>
            Extract {index + 1} / {total}
          </Text>
          <Text variant="xSmall" styles={{ root: { color: '#a19f9d' } }}>
            · {charCount.toLocaleString()} chars
          </Text>
        </span>
        {typeof extract.relevanceScore === 'number' && (
          <ScoreChip score={extract.relevanceScore} />
        )}
      </button>
      {expanded && (
        <div
          id={panelId}
          style={{
            padding: '8px 10px 10px 10px',
            fontSize: 13,
            lineHeight: 1.5,
            color: '#323130',
            borderTop: '1px solid #edebe9'
          }}
        >
          <div
            style={{
              display: 'inline-flex',
              gap: 0,
              marginBottom: 8,
              borderRadius: 3,
              overflow: 'hidden'
            }}
            role="tablist"
            aria-label="Extract view mode"
          >
            <button
              type="button"
              role="tab"
              aria-selected={mode === 'formatted'}
              onClick={(e) => {
                e.stopPropagation();
                setMode('formatted');
              }}
              style={{ ...toggleBtnStyle(mode === 'formatted'), borderRadius: '3px 0 0 3px' }}
            >
              Formatted
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={mode === 'raw'}
              onClick={(e) => {
                e.stopPropagation();
                setMode('raw');
              }}
              style={{
                ...toggleBtnStyle(mode === 'raw'),
                borderRadius: '0 3px 3px 0',
                borderLeft: 'none'
              }}
            >
              Raw
            </button>
          </div>
          {mode === 'formatted' ? (
            <div>{renderMiniMarkdown(formattedSource)}</div>
          ) : (
            <pre
              style={{
                margin: 0,
                padding: 10,
                background: '#ffffff',
                border: '1px solid #edebe9',
                borderRadius: 3,
                fontFamily: 'Consolas, "Courier New", monospace',
                fontSize: 12,
                lineHeight: 1.5,
                whiteSpace: 'pre-wrap',
                wordBreak: 'break-word',
                maxHeight: 360,
                overflow: 'auto'
              }}
            >
              {extract.text}
            </pre>
          )}
        </div>
      )}
    </div>
  );
}

/** Normalise the raw extract text returned by the Retrieval API.
 *  - Collapse \r\n to \n.
 *  - Strip a few noisy WebPart JSON blobs that occasionally leak through.
 *  - Collapse 3+ consecutive blank lines down to 2.
 */
function cleanExtractText(raw: string): string {
  let text = raw.replace(/\r\n/g, '\n');
  // Drop obvious WebPart JSON blobs (they start with `{ "webPartData":` ...).
  text = text.replace(/\{\s*"webPartData"[\s\S]*?\}\s*\}\s*\}\s*\}/g, '');
  text = text.replace(/\n{3,}/g, '\n\n');
  return text.trim();
}

/** Tiny inline markdown renderer (no innerHTML, no third-party lib). Handles:
 *  - ATX headings (#### Heading)
 *  - Unordered lists (- item)
 *  - Inline links [text](url)
 *  - Plain paragraphs (blank line separated)
 */
function renderMiniMarkdown(text: string): React.ReactNode {
  const lines = text.split('\n');
  const nodes: React.ReactNode[] = [];
  let listBuffer: string[] = [];
  let key = 0;

  const flushList = (): void => {
    if (listBuffer.length === 0) {
      return;
    }
    nodes.push(
      <ul key={key++} style={{ margin: '4px 0 4px 18px', padding: 0 }}>
        {listBuffer.map((item, i) => (
          <li key={i} style={{ marginBottom: 2 }}>
            {renderInline(item)}
          </li>
        ))}
      </ul>
    );
    listBuffer = [];
  };

  for (let i = 0; i < lines.length; i++) {
    const rawLine = lines[i];
    const line = rawLine.replace(/\s+$/, '');

    const listMatch = /^\s*-\s+(.*)$/.exec(line);
    if (listMatch) {
      listBuffer.push(listMatch[1]);
      continue;
    }
    flushList();

    const headingMatch = /^(#{1,6})\s+(.*)$/.exec(line);
    if (headingMatch) {
      const level = headingMatch[1].length;
      const content = headingMatch[2];
      const fontSize = level <= 2 ? 15 : level <= 4 ? 13 : 12;
      nodes.push(
        <div
          key={key++}
          style={{
            fontWeight: 700,
            fontSize,
            marginTop: 8,
            marginBottom: 2,
            color: '#201f1e'
          }}
        >
          {renderInline(content)}
        </div>
      );
      continue;
    }

    if (line.length === 0) {
      nodes.push(<div key={key++} style={{ height: 6 }} />);
      continue;
    }

    nodes.push(
      <div key={key++} style={{ margin: '2px 0' }}>
        {renderInline(line)}
      </div>
    );
  }
  flushList();
  return nodes;
}

/** Inline renderer with light support for the markup that ships inside
 *  Retrieval extracts:
 *  - [text](url) Markdown links
 *  - <span style="color:#XXX">...</span> coloured runs (other style props
 *    are ignored for safety)
 *  - **bold** and __bold__
 *  - _italic_ and *italic*
 *  Anything that doesn't match is rendered verbatim. No innerHTML is used. */
function renderInline(text: string): React.ReactNode {
  return parseInline(text, 0);
}

let inlineKey = 0;

function parseInline(text: string, depth: number): React.ReactNode {
  if (depth > 4 || text.length === 0) {
    return text;
  }
  const parts: React.ReactNode[] = [];
  let i = 0;
  let buf = '';
  const flush = (): void => {
    if (buf.length > 0) {
      parts.push(buf);
      buf = '';
    }
  };
  while (i < text.length) {
    const rest = text.substring(i);

    // <span style="color:#XXX">...</span>
    const spanOpen = /^<span\b[^>]*style\s*=\s*"([^"]*)"[^>]*>/i.exec(rest);
    if (spanOpen) {
      const closeIdx = rest.indexOf('</span>', spanOpen[0].length);
      if (closeIdx !== -1) {
        flush();
        const inner = rest.substring(spanOpen[0].length, closeIdx);
        const colorMatch = /color\s*:\s*(#[0-9a-fA-F]{3,8}|[a-zA-Z]+)/.exec(spanOpen[1]);
        const color = colorMatch ? colorMatch[1] : undefined;
        parts.push(
          <span key={inlineKey++} style={color ? { color } : undefined}>
            {parseInline(inner, depth + 1)}
          </span>
        );
        i += closeIdx + '</span>'.length;
        continue;
      }
    }

    // Markdown link [text](url)
    if (rest.charCodeAt(0) === 91 /* [ */) {
      const linkMatch = /^\[([^\]]+)\]\(([^)]+)\)/.exec(rest);
      if (linkMatch) {
        flush();
        parts.push(
          <a
            key={inlineKey++}
            href={linkMatch[2]}
            target="_blank"
            rel="noreferrer"
            style={{ wordBreak: 'break-all' }}
          >
            {parseInline(linkMatch[1], depth + 1)}
          </a>
        );
        i += linkMatch[0].length;
        continue;
      }
    }

    // **bold** or __bold__
    if ((rest.startsWith('**') || rest.startsWith('__')) && rest.length > 2) {
      const marker = rest.substring(0, 2);
      const endIdx = rest.indexOf(marker, 2);
      if (endIdx > 2) {
        flush();
        parts.push(
          <strong key={inlineKey++}>{parseInline(rest.substring(2, endIdx), depth + 1)}</strong>
        );
        i += endIdx + 2;
        continue;
      }
    }

    // _italic_ or *italic* (single-char markers, avoid matching inside words)
    const c = rest.charAt(0);
    if ((c === '_' || c === '*') && rest.length > 1) {
      // Don't treat ** or __ as italic (already handled above).
      if (rest.charAt(1) !== c) {
        const prevChar = i > 0 ? text.charAt(i - 1) : ' ';
        if (!/[A-Za-z0-9]/.test(prevChar)) {
          const endRegex = c === '_' ? /[^A-Za-z0-9]_/ : /[^A-Za-z0-9]\*/;
          const m = endRegex.exec(rest.substring(1));
          if (m) {
            const endIdx = m.index + 1; // position of marker inside rest
            const inner = rest.substring(1, endIdx);
            if (inner.length > 0 && !inner.includes('\n')) {
              flush();
              parts.push(
                <em key={inlineKey++}>{parseInline(inner, depth + 1)}</em>
              );
              i += endIdx + 1;
              // Re-emit the non-word boundary char we matched as the closer.
              // The regex consumed one char before the marker; we need to keep it.
              // m.index points to the char before the marker in rest.substring(1),
              // so the marker char itself is at endIdx in rest. We already advanced
              // past it; the boundary char (rest[endIdx-? ]) is part of inner if any.
              // Simpler: nothing to re-emit because boundary char was inside inner span.
              continue;
            }
          }
        }
      }
    }

    buf += text.charAt(i);
    i++;
  }
  flush();
  return parts;
}

/** Prepares a cleaned extract for the Formatted view by translating the
 *  `<slide_N>` / `</slide_N>` pseudo-tags into a visible heading + divider so
 *  the structure of slide decks survives the rendering. */
function prepareFormattedExtract(text: string): string {
  let out = text;
  out = out.replace(/<slide_(\d+)>/gi, (_m, n) => `\n\n##### Slide ${n}\n`);
  out = out.replace(/<\/slide_\d+>/gi, '\n');
  out = out.replace(/\n{3,}/g, '\n\n');
  return out.trim();
}

function renderInteractions(body: unknown): React.ReactNode {
  const items = (body as { value?: IAiInteraction[] })?.value;
  if (!items || items.length === 0) {
    return <Text variant="small">No interactions returned.</Text>;
  }
  return <InteractionsBySession items={items} />;
}

const NO_SESSION_KEY = '__no_session__';

interface ISessionGroup {
  sessionId: string;
  /** Display label: the real session id or "(no sessionId)". */
  label: string;
  items: IAiInteraction[];
  latest: string; // ISO date of the most recent interaction in the group
  appClasses: Set<string>;
}

function groupInteractionsBySession(items: IAiInteraction[]): ISessionGroup[] {
  const map = new Map<string, ISessionGroup>();
  for (const item of items) {
    const key = item.sessionId ?? NO_SESSION_KEY;
    let group = map.get(key);
    if (!group) {
      group = {
        sessionId: key,
        label: item.sessionId ?? '(no sessionId)',
        items: [],
        latest: '',
        appClasses: new Set<string>()
      };
      map.set(key, group);
    }
    group.items.push(item);
    if (item.appClass) {
      group.appClasses.add(item.appClass);
    }
    const dt = item.createdDateTime ?? '';
    if (dt > group.latest) {
      group.latest = dt;
    }
  }
  // Sort sessions newest-first by their most recent interaction.
  return Array.from(map.values()).sort((a, b) => (a.latest < b.latest ? 1 : a.latest > b.latest ? -1 : 0));
}

function InteractionsBySession(props: { items: IAiInteraction[] }): JSX.Element {
  const { items } = props;
  const groups = React.useMemo(() => groupInteractionsBySession(items), [items]);
  // First session expanded by default; rest collapsed. Keyed by sessionId.
  const [expandedSessions, setExpandedSessions] = React.useState<Record<string, boolean>>(() => {
    const init: Record<string, boolean> = {};
    groups.forEach((g, i) => {
      init[g.sessionId] = i === 0;
    });
    return init;
  });

  const toggleSession = (sid: string): void => {
    setExpandedSessions((prev) => ({ ...prev, [sid]: !prev[sid] }));
  };

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <Text variant="xSmall" styles={{ root: { color: '#605e5c' } }}>
        {items.length} {items.length === 1 ? 'interaction' : 'interactions'} grouped into{' '}
        {groups.length} {groups.length === 1 ? 'session' : 'sessions'} · sessions sorted by most recent
        activity.
      </Text>
      {groups.map((group) => (
        <SessionCard
          key={group.sessionId}
          group={group}
          expanded={!!expandedSessions[group.sessionId]}
          onToggle={() => toggleSession(group.sessionId)}
        />
      ))}
    </Stack>
  );
}

function SessionCard(props: { group: ISessionGroup; expanded: boolean; onToggle: () => void }): JSX.Element {
  const { group, expanded, onToggle } = props;
  // Conversation order: oldest first, so we read the thread top-to-bottom like
  // a real chat (prompt -> response -> follow-up prompt -> response).
  const ordered = React.useMemo(() => {
    return [...group.items].sort((a, b) => {
      const ad = a.createdDateTime ?? '';
      const bd = b.createdDateTime ?? '';
      return ad < bd ? -1 : ad > bd ? 1 : 0;
    });
  }, [group.items]);

  const counts = React.useMemo(() => {
    const c: Record<string, number> = {};
    for (const it of group.items) {
      const t = it.interactionType ?? '(unknown)';
      c[t] = (c[t] ?? 0) + 1;
    }
    return c;
  }, [group.items]);

  const panelId = `session-${group.sessionId}`;

  return (
    <div
      style={{
        border: '1px solid #edebe9',
        borderRadius: 4,
        background: '#ffffff',
        overflow: 'hidden'
      }}
    >
      <button
        type="button"
        onClick={onToggle}
        aria-expanded={expanded}
        aria-controls={panelId}
        style={{
          width: '100%',
          background: '#f3f2f1',
          border: 'none',
          padding: '8px 12px',
          cursor: 'pointer',
          textAlign: 'left',
          display: 'flex',
          alignItems: 'center',
          gap: 10
        }}
      >
        <Icon
          iconName="ChevronRight"
          style={{
            fontSize: 10,
            color: '#605e5c',
            transform: expanded ? 'rotate(90deg)' : 'rotate(0deg)',
            transition: 'transform 0.15s ease'
          }}
        />
        <div style={{ flex: 1, minWidth: 0 }}>
          <div
            style={{
              fontFamily: 'Consolas, monospace',
              fontSize: 12,
              fontWeight: 600,
              color: '#201f1e',
              overflow: 'hidden',
              textOverflow: 'ellipsis',
              whiteSpace: 'nowrap'
            }}
            title={group.label}
          >
            <Icon iconName="ChatBubbles" /> {group.label}
          </div>
          <Text variant="xSmall" styles={{ root: { color: '#605e5c' } }}>
            {group.items.length} {group.items.length === 1 ? 'interaction' : 'interactions'}
            {counts.userPrompt ? ` · ${counts.userPrompt} prompt${counts.userPrompt === 1 ? '' : 's'}` : ''}
            {counts.aiResponse ? ` · ${counts.aiResponse} response${counts.aiResponse === 1 ? '' : 's'}` : ''}
            {group.latest ? ` · latest ${group.latest}` : ''}
            {group.appClasses.size > 0 ? ` · ${Array.from(group.appClasses).join(', ')}` : ''}
          </Text>
        </div>
      </button>
      {expanded && (
        <div
          id={panelId}
          style={{ padding: 12, borderTop: '1px solid #edebe9', background: '#fafafa' }}
        >
          <Stack tokens={{ childrenGap: 10 }}>
            {ordered.map((item) => (
              <ChatBubble key={item.id} item={item} />
            ))}
          </Stack>
        </div>
      )}
    </div>
  );
}

/** Extract the most useful display text out of an interaction's body.
 *  - For html bodies that are just an <attachment id="X"> placeholder, pulls
 *    the AdaptiveCard text block referenced by that attachment (this is how
 *    Copilot's real answers ship).
 *  - Otherwise returns the raw body content. */
function extractInteractionText(item: IAiInteraction): string {
  const content = item.body?.content ?? '';
  const isHtml = item.body?.contentType === 'html';
  if (!isHtml) {
    return content;
  }
  const attMatch = /<attachment\s+id=\"([^\"]+)\"\s*>/i.exec(content);
  if (!attMatch || !item.attachments) {
    return content;
  }
  const att = item.attachments.find((a) => a.attachmentId === attMatch[1]);
  if (!att || !att.content) {
    return content;
  }
  try {
    const card = JSON.parse(att.content) as {
      body?: Array<{ type?: string; text?: string }>;
    };
    if (Array.isArray(card.body)) {
      const texts = card.body
        .filter((b) => b && b.type === 'TextBlock' && typeof b.text === 'string')
        .map((b) => b.text as string);
      if (texts.length > 0) {
        return texts.join('\n\n');
      }
    }
  } catch {
    // Not an adaptive card – fall through.
  }
  return att.content;
}

function ChatBubble(props: { item: IAiInteraction }): JSX.Element {
  const { item } = props;
  const isPrompt = item.interactionType === 'userPrompt';
  const isResponse = item.interactionType === 'aiResponse';
  const bg = isPrompt ? '#fff4ce' : isResponse ? '#deecf9' : '#f3f2f1';
  const fg = isPrompt ? '#8a6d00' : isResponse ? '#004578' : '#605e5c';
  const iconName = isPrompt ? 'Contact' : isResponse ? 'Robot' : 'Tag';
  const senderName =
    item.from?.user?.displayName ||
    item.from?.application?.displayName ||
    (isPrompt ? 'User' : isResponse ? 'Assistant' : item.interactionType ?? 'message');

  const [showAll, setShowAll] = React.useState<boolean>(false);
  const text = extractInteractionText(item);
  const isLong = text.length > 600;
  const visible = !isLong || showAll ? text : `${text.substring(0, 600)}…`;

  return (
    <div
      style={{
        display: 'flex',
        justifyContent: isPrompt ? 'flex-end' : 'flex-start'
      }}
    >
      <div
        style={{
          maxWidth: '85%',
          background: '#ffffff',
          border: '1px solid #edebe9',
          borderLeft: `3px solid ${fg}`,
          borderRadius: 6,
          padding: '8px 12px',
          boxShadow: '0 1px 0 rgba(0,0,0,0.02)'
        }}
      >
        <Stack
          horizontal
          verticalAlign="center"
          tokens={{ childrenGap: 6 }}
          styles={{ root: { marginBottom: 4 } }}
        >
          <span
            style={{
              display: 'inline-flex',
              alignItems: 'center',
              gap: 4,
              background: bg,
              color: fg,
              borderRadius: 10,
              padding: '1px 8px',
              fontSize: 11,
              fontWeight: 700
            }}
          >
            <Icon iconName={iconName} /> {item.interactionType ?? 'message'}
          </span>
          <Text variant="xSmall" styles={{ root: { color: '#605e5c', flexGrow: 1 } }}>
            {senderName}
            {item.createdDateTime ? ` · ${item.createdDateTime}` : ''}
          </Text>
        </Stack>
        <div
          style={{
            fontSize: 13,
            lineHeight: 1.5,
            whiteSpace: 'pre-wrap',
            wordBreak: 'break-word',
            color: '#201f1e'
          }}
        >
          {visible || '(no text content)'}
        </div>
        {isLong && (
          <button
            type="button"
            onClick={() => setShowAll((v) => !v)}
            style={{
              marginTop: 6,
              background: 'transparent',
              border: 'none',
              color: '#0078d4',
              cursor: 'pointer',
              padding: 0,
              fontSize: 12,
              fontWeight: 600
            }}
          >
            {showAll ? 'Show less' : `Show full message (${text.length.toLocaleString()} chars)`}
          </button>
        )}
      </div>
    </div>
  );
}

// -----------------------------------------------------------------------------
// Copilot Usage report renderers
// -----------------------------------------------------------------------------
// All three usage endpoints return rich, structured payloads that we know the
// shape of (see Microsoft 365 Copilot adoption reports). We render each one
// with a purpose-built view that highlights what matters on stage.
// -----------------------------------------------------------------------------

/** The Copilot host apps that the adoption reports break out, in display order. */
const COPILOT_APPS: { key: string; label: string }[] = [
  { key: 'copilotChat', label: 'Copilot Chat' },
  { key: 'microsoftTeams', label: 'Teams' },
  { key: 'word', label: 'Word' },
  { key: 'excel', label: 'Excel' },
  { key: 'powerPoint', label: 'PowerPoint' },
  { key: 'outlook', label: 'Outlook' },
  { key: 'oneNote', label: 'OneNote' },
  { key: 'loop', label: 'Loop' }
];

/** Shared palette for the bar chart. */
const BAR_TRACK = '#f3f2f1';
const BAR_ACTIVE = '#0078d4';
const BAR_HIGHLIGHT = '#107c10';

/** Header strip with refresh-date + secondary metadata. */
const ReportHeader: React.FC<{ refreshDate?: string; extras?: string[] }> = ({
  refreshDate,
  extras
}) => (
  <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center" wrap>
    {refreshDate && (
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
        <Icon iconName="Refresh" /> Refresh date: {refreshDate}
      </span>
    )}
    {(extras ?? []).map((text, i) => (
      <span
        key={i}
        style={{
          background: '#f3f2f1',
          color: '#323130',
          borderRadius: 12,
          padding: '2px 10px',
          fontSize: 12,
          fontWeight: 600
        }}
      >
        {text}
      </span>
    ))}
  </Stack>
);

// ---- 1. Usage Summary ------------------------------------------------------

interface IAdoptionRow {
  reportPeriod?: number;
  [key: string]: number | undefined;
}

function renderUsageSummary(body: unknown): React.ReactNode {
  const top = (body as { value?: { reportRefreshDate?: string; adoptionByProduct?: IAdoptionRow[] }[] })
    ?.value?.[0];
  if (!top) {
    return <Text variant="small">No summary data returned.</Text>;
  }
  const rows = top.adoptionByProduct ?? [];
  if (rows.length === 0) {
    return <Text variant="small">The response has no adoptionByProduct entries.</Text>;
  }

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <ReportHeader refreshDate={top.reportRefreshDate} />
      {rows.map((row, rowIdx) => {
        // anyApp is the headline metric — show it as a large stat card on top.
        const anyEnabled = row.anyAppEnabledUsers ?? 0;
        const anyActive = row.anyAppActiveUsers ?? 0;
        const anyRate = anyEnabled > 0 ? Math.round((anyActive / anyEnabled) * 100) : 0;
        return (
          <Stack key={rowIdx} tokens={{ childrenGap: 12 }}>
            <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
              <Icon iconName="Calendar" /> Last {row.reportPeriod ?? '?'} days
            </Text>
            <div
              style={{
                display: 'grid',
                gridTemplateColumns: 'repeat(3, minmax(0, 1fr))',
                gap: 12
              }}
            >
              <StatCard label="Any-app active" value={anyActive} sub={`of ${anyEnabled} licensed`} accent />
              <StatCard label="Adoption rate" value={`${anyRate}%`} sub="active / licensed" />
              <StatCard label="Period" value={`D${row.reportPeriod ?? '?'}`} sub="report window" />
            </div>
            <Text variant="small" styles={{ root: { fontWeight: 600, marginTop: 4 } }}>
              Active users per Copilot host (active / enabled)
            </Text>
            <Stack tokens={{ childrenGap: 6 }}>
              {COPILOT_APPS.map((app) => {
                const enabled = row[`${app.key}EnabledUsers`] ?? 0;
                const active = row[`${app.key}ActiveUsers`] ?? 0;
                return (
                  <AppBar
                    key={app.key}
                    label={app.label}
                    active={active}
                    enabled={enabled}
                    highlight={app.key === 'copilotChat'}
                  />
                );
              })}
            </Stack>
          </Stack>
        );
      })}
    </Stack>
  );
}

/** Small KPI card used by the Summary header. */
function StatCard(props: { label: string; value: number | string; sub?: string; accent?: boolean }): JSX.Element {
  const { label, value, sub, accent } = props;
  return (
  <div
    style={{
      border: `1px solid ${accent ? '#0078d4' : '#edebe9'}`,
      borderLeft: `4px solid ${accent ? '#0078d4' : '#605e5c'}`,
      borderRadius: 4,
      padding: 12,
      background: '#fff'
    }}
  >
    <Text variant="xSmall" styles={{ root: { color: '#605e5c', textTransform: 'uppercase' } }}>
      {label}
    </Text>
    <Text variant="xxLarge" styles={{ root: { fontWeight: 700, color: accent ? '#0078d4' : '#323130' } }}>
      {value}
    </Text>
    {sub && (
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        {sub}
      </Text>
    )}
  </div>
  );
}

/** A single active/enabled horizontal bar with rate label. */
function AppBar(props: { label: string; active: number; enabled: number; highlight?: boolean }): JSX.Element {
  const { label, active, enabled, highlight } = props;
  const pct = enabled > 0 ? (active / enabled) * 100 : 0;
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
      <span
        style={{
          width: 110,
          fontSize: 12,
          fontWeight: highlight ? 600 : 400,
          color: highlight ? '#004578' : '#323130'
        }}
        title={label}
      >
        {label}
      </span>
      <div style={{ flex: 1, background: BAR_TRACK, borderRadius: 4, height: 18, position: 'relative' }}>
        <div
          style={{
            width: `${Math.max(active > 0 ? 2 : 0, pct)}%`,
            background: highlight ? BAR_HIGHLIGHT : BAR_ACTIVE,
            height: '100%',
            borderRadius: 4
          }}
        />
      </div>
      <span style={{ width: 110, textAlign: 'right', fontSize: 12, fontWeight: 600 }}>
        {active} / {enabled}
        <span style={{ color: '#605e5c', fontWeight: 400 }}> ({Math.round(pct)}%)</span>
      </span>
    </div>
  );
}

// ---- 2. Usage Trend (interactive) ------------------------------------------

interface ITrendDay {
  reportDate?: string;
  [key: string]: number | string | undefined;
}

/** Interactive trend chart: select an app and we plot its daily active-users bars. */
const UsageTrendChart: React.FC<{ body: unknown }> = ({ body }) => {
  const top = (body as {
    value?: { reportRefreshDate?: string; reportPeriod?: number; adoptionByDate?: ITrendDay[] }[];
  })?.value?.[0];
  const days = top?.adoptionByDate ?? [];
  const [selectedApp, setSelectedApp] = React.useState<string>('copilotChat');

  if (days.length === 0) {
    return <Text variant="small">No adoptionByDate entries in the response.</Text>;
  }

  // Sort ascending by date so the chart reads left-to-right (oldest -> newest).
  const sorted = [...days].sort((a, b) => (a.reportDate ?? '').localeCompare(b.reportDate ?? ''));
  const enabledForApp = Number(sorted[sorted.length - 1]?.[`${selectedApp}EnabledUsers`] ?? 0);
  const series = sorted.map((d) => ({
    date: d.reportDate ?? '',
    active: Number(d[`${selectedApp}ActiveUsers`] ?? 0)
  }));
  const max = Math.max(enabledForApp, ...series.map((s) => s.active), 1);
  const totalActive = series.reduce((sum, s) => sum + s.active, 0);
  const daysWithActivity = series.filter((s) => s.active > 0).length;

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <ReportHeader
        refreshDate={top?.reportRefreshDate}
        extras={[`Period: D${top?.reportPeriod ?? '?'}`, `${days.length} days returned`]}
      />

      {/* App selector */}
      <Stack horizontal tokens={{ childrenGap: 6 }} wrap>
        {COPILOT_APPS.map((app) => {
          const isSelected = app.key === selectedApp;
          return (
            <button
              key={app.key}
              type="button"
              onClick={() => setSelectedApp(app.key)}
              style={{
                border: `1px solid ${isSelected ? '#0078d4' : '#edebe9'}`,
                background: isSelected ? '#0078d4' : '#fff',
                color: isSelected ? '#fff' : '#323130',
                borderRadius: 14,
                padding: '4px 10px',
                fontSize: 12,
                fontWeight: 600,
                cursor: 'pointer'
              }}
            >
              {app.label}
            </button>
          );
        })}
      </Stack>

      {/* Summary line for the selected app */}
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Showing daily <strong>active users</strong> for <strong>{COPILOT_APPS.find((a) => a.key === selectedApp)?.label}</strong>
        {' · '}
        {enabledForApp} licensed users · {daysWithActivity}/{series.length} days with activity · {totalActive} total active-user-days
      </Text>

      {/* Vertical column chart */}
      <div
        style={{
          display: 'flex',
          alignItems: 'flex-end',
          gap: 2,
          height: 160,
          padding: '8px 4px',
          background: '#fafafa',
          border: '1px solid #edebe9',
          borderRadius: 4,
          overflowX: 'auto'
        }}
      >
        {series.map((s) => {
          const h = (s.active / max) * 100;
          const isToday = s.date === top?.reportRefreshDate;
          return (
            <div
              key={s.date}
              title={`${s.date}: ${s.active} active`}
              style={{
                display: 'flex',
                flexDirection: 'column',
                justifyContent: 'flex-end',
                alignItems: 'center',
                minWidth: 18,
                flex: '0 0 auto',
                height: '100%'
              }}
            >
              <span
                style={{
                  fontSize: 10,
                  color: s.active > 0 ? '#323130' : 'transparent',
                  marginBottom: 2,
                  fontWeight: 600
                }}
              >
                {s.active}
              </span>
              <div
                style={{
                  width: '100%',
                  height: `${Math.max(s.active > 0 ? 4 : 0, h)}%`,
                  background: isToday ? BAR_HIGHLIGHT : BAR_ACTIVE,
                  borderRadius: '2px 2px 0 0'
                }}
              />
            </div>
          );
        })}
      </div>

      {/* X-axis labels (compact: only show every Nth date) */}
      <div style={{ display: 'flex', gap: 2, paddingLeft: 4, overflowX: 'auto' }}>
        {series.map((s, i) => {
          const showLabel = i === 0 || i === series.length - 1 || i % Math.ceil(series.length / 6) === 0;
          return (
            <div
              key={s.date}
              style={{
                minWidth: 18,
                flex: '0 0 auto',
                fontSize: 9,
                color: '#605e5c',
                textAlign: 'center',
                writingMode: 'vertical-rl',
                transform: 'rotate(180deg)',
                height: 56
              }}
            >
              {showLabel ? s.date : ''}
            </div>
          );
        })}
      </div>
    </Stack>
  );
};

function renderUsageTrend(body: unknown): React.ReactNode {
  return <UsageTrendChart body={body} />;
}

// ---- 3. Usage User Detail --------------------------------------------------

interface IUsageUser {
  userPrincipalName?: string;
  displayName?: string;
  lastActivityDate?: string;
  copilotChatLastActivityDate?: string;
  microsoftTeamsCopilotLastActivityDate?: string;
  wordCopilotLastActivityDate?: string;
  excelCopilotLastActivityDate?: string;
  powerPointCopilotLastActivityDate?: string;
  outlookCopilotLastActivityDate?: string;
  oneNoteCopilotLastActivityDate?: string;
  loopCopilotLastActivityDate?: string;
  reportRefreshDate?: string;
}

/** Pick the lastActivityDate field for a given app key from a per-user row. */
function userAppDate(user: IUsageUser, appKey: string): string {
  // chat field is named `copilotChatLastActivityDate` (no "Copilot" infix),
  // all the others follow `<app>CopilotLastActivityDate`.
  const fieldName =
    appKey === 'copilotChat' ? 'copilotChatLastActivityDate' : `${appKey}CopilotLastActivityDate`;
  return ((user as unknown) as Record<string, string>)[fieldName] ?? '';
}

/** Colour a date cell based on recency vs the report's refresh date. */
function recencyStyle(dateStr: string, refreshDate: string): { bg: string; color: string } {
  if (!dateStr) {
    return { bg: '#faf9f8', color: '#a19f9d' };
  }
  const ref = refreshDate ? new Date(refreshDate).getTime() : Date.now();
  const d = new Date(dateStr).getTime();
  if (isNaN(d)) {
    return { bg: '#faf9f8', color: '#a19f9d' };
  }
  const days = Math.floor((ref - d) / (1000 * 60 * 60 * 24));
  if (days <= 7) {
    return { bg: '#dff6dd', color: '#0b6a0b' };
  }
  if (days <= 30) {
    return { bg: '#e8f5e8', color: '#107c10' };
  }
  if (days <= 90) {
    return { bg: '#fff4ce', color: '#8a6d00' };
  }
  return { bg: '#fde7e9', color: '#a4262c' };
}

function renderUsageUserDetail(body: unknown): React.ReactNode {
  const users = ((body as { value?: IUsageUser[] })?.value ?? []).slice();
  if (users.length === 0) {
    return <Text variant="small">No user rows returned.</Text>;
  }
  const refreshDate = users[0]?.reportRefreshDate ?? '';

  // Sort users by their overall lastActivityDate (most recent first); empty dates last.
  users.sort((a, b) => (b.lastActivityDate ?? '').localeCompare(a.lastActivityDate ?? ''));

  // KPIs
  const totalUsers = users.length;
  const activeUsers = users.filter((u) => !!u.lastActivityDate).length;
  const chatUsers = users.filter((u) => !!u.copilotChatLastActivityDate).length;

  const cell: React.CSSProperties = {
    border: '1px solid #edebe9',
    padding: '6px 8px',
    fontSize: 11,
    verticalAlign: 'middle',
    whiteSpace: 'nowrap'
  };
  const headerCell: React.CSSProperties = { ...cell, background: '#f3f2f1', fontWeight: 600 };

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <ReportHeader refreshDate={refreshDate} />
      <div
        style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(3, minmax(0, 1fr))',
          gap: 12
        }}
      >
        <StatCard label="Users in report" value={totalUsers} />
        <StatCard label="With any activity" value={activeUsers} sub={`${Math.round((activeUsers / totalUsers) * 100)}%`} accent />
        <StatCard label="Copilot Chat users" value={chatUsers} sub={`${Math.round((chatUsers / totalUsers) * 100)}%`} />
      </div>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Each cell shows the user&apos;s last activity date in that surface, coloured by recency relative to the refresh date
        (green ≤30 days, amber ≤90 days, red &gt;90 days, grey if no activity).
      </Text>
      <div style={{ overflowX: 'auto' }}>
        <table style={{ borderCollapse: 'collapse', width: '100%' }}>
          <thead>
            <tr>
              <th style={headerCell}>User</th>
              <th style={headerCell}>Last activity</th>
              {COPILOT_APPS.map((app) => (
                <th key={app.key} style={headerCell}>
                  {app.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {users.map((u, i) => {
              const overall = recencyStyle(u.lastActivityDate ?? '', refreshDate);
              return (
                <tr key={i}>
                  <td style={cell}>
                    <div style={{ fontWeight: 600 }}>{u.displayName ?? '(no name)'}</div>
                    <div style={{ color: '#605e5c', fontSize: 10 }}>{u.userPrincipalName ?? ''}</div>
                  </td>
                  <td style={{ ...cell, background: overall.bg, color: overall.color, fontWeight: 600 }}>
                    {u.lastActivityDate || '—'}
                  </td>
                  {COPILOT_APPS.map((app) => {
                    const d = userAppDate(u, app.key);
                    const s = recencyStyle(d, refreshDate);
                    return (
                      <td
                        key={app.key}
                        style={{ ...cell, background: s.bg, color: s.color, fontWeight: d ? 600 : 400 }}
                      >
                        {d || '—'}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </Stack>
  );
}

const RICH_RENDERERS: Record<ApiId, (body: unknown) => React.ReactNode> = {
  retrieval: renderRetrieval,
  interactionExport: renderInteractions,
  usageSummary: renderUsageSummary,
  usageTrend: renderUsageTrend,
  usageUserDetail: renderUsageUserDetail
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

            {/* Raw, syntax-highlighted JSON with a line-number gutter
                and a click-to-expand fullscreen modal. */}
            <RawJsonViewer value={result.body} />
          </Stack>
        )}
      </Stack>
    </div>
  );
};

export default ResponseViewer;
