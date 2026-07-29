import * as React from 'react';
import {
  DetailsList,
  IColumn,
  SelectionMode,
  Link,
  IGroup,
  IGroupRenderProps,
  IGroupHeaderProps,
  Icon,
  Text,
  useTheme,
  mergeStyles
} from '@fluentui/react';

import { IDriftItem } from '../../../../models/IDriftItem';
import { DriftStatus } from '../../../../models/DriftStatus';
import StatusPill from './StatusPill';

export interface IDriftTableProps {
  items: IDriftItem[];
}

type SortableField = 'Title' | 'DriftStatus' | 'DaysDrift';

interface ISortState {
  field: SortableField;
  isDescending: boolean;
}

const DEFAULT_SORT_STATE: ISortState = { field: 'DaysDrift', isDescending: true };

function formatDate(value: string | undefined): string {
  if (!value) {
    return '—';
  }
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? '—' : parsed.toLocaleDateString();
}

function getSiteName(siteUrl: string): string {
  const trimmed = siteUrl.replace(/\/$/, '');
  const lastSlash = trimmed.lastIndexOf('/');
  return lastSlash >= 0 ? trimmed.substring(lastSlash + 1) : siteUrl;
}

// Primary sort: SiteUrl (groups contiguous); secondary: user's chosen column.
function sortItemsGrouped(items: IDriftItem[], sortState: ISortState): IDriftItem[] {
  return [...items].sort((a, b) => {
    const siteCompare = a.SiteUrl.localeCompare(b.SiteUrl);
    if (siteCompare !== 0) return siteCompare;

    let comparison = 0;
    if (sortState.field === 'DaysDrift') {
      comparison = a.DaysDrift - b.DaysDrift;
    } else if (sortState.field === 'Title') {
      comparison = a.Title.localeCompare(b.Title);
    } else {
      comparison = a.DriftStatus.localeCompare(b.DriftStatus);
    }
    return sortState.isDescending ? -comparison : comparison;
  });
}

function buildGroups(sortedItems: IDriftItem[]): IGroup[] {
  const groups: IGroup[] = [];

  for (let i = 0; i < sortedItems.length; i++) {
    const siteUrl = sortedItems[i].SiteUrl;
    if (groups.length === 0 || groups[groups.length - 1].key !== siteUrl) {
      groups.push({
        key: siteUrl,
        name: getSiteName(siteUrl),
        startIndex: i,
        count: 1,
        isCollapsed: false,
        level: 0
      });
    } else {
      groups[groups.length - 1].count++;
    }
  }

  return groups;
}

// ---------------------------------------------------------------------------
// Custom group header component
// ---------------------------------------------------------------------------

interface IDriftGroupHeaderComponentProps {
  groupHeaderProps: IGroupHeaderProps;
  allItems: IDriftItem[];
}

function DriftGroupHeader(componentProps: IDriftGroupHeaderComponentProps): JSX.Element {
  const { groupHeaderProps, allItems } = componentProps;
  const theme = useTheme();
  const group = groupHeaderProps.group!;

  const groupItems = allItems.slice(group.startIndex, group.startIndex + group.count);
  const staleCount = groupItems.filter((i) => i.DriftStatus === DriftStatus.Stale).length;
  const missingCount = groupItems.filter((i) => i.DriftStatus === DriftStatus.Missing).length;
  const abandonedCount = groupItems.filter((i) => i.DriftStatus === DriftStatus.Abandoned).length;

  const isCollapsed = group.isCollapsed ?? false;

  const handleToggle = (): void => {
    if (groupHeaderProps.onToggleCollapse) {
      groupHeaderProps.onToggleCollapse(group);
    }
  };

  const headerClass = mergeStyles({
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    padding: '8px 12px',
    backgroundColor: theme.palette.neutralLighterAlt,
    borderBottom: `1px solid ${theme.palette.neutralLight}`,
    cursor: 'pointer',
    userSelect: 'none',
    selectors: {
      ':hover': {
        backgroundColor: theme.palette.neutralLighter
      }
    }
  });

  const pillStyle = (bg: string, color: string): React.CSSProperties => ({
    display: 'inline-block',
    padding: '1px 8px',
    borderRadius: '10px',
    fontSize: '11px',
    fontWeight: 600,
    backgroundColor: bg,
    color
  });

  return (
    <div
      className={headerClass}
      role="button"
      tabIndex={0}
      onClick={handleToggle}
      onKeyDown={(e: React.KeyboardEvent<HTMLDivElement>): void => {
        if (e.key === 'Enter' || e.key === ' ') {
          handleToggle();
        }
      }}
    >
      <Icon
        iconName={isCollapsed ? 'ChevronRight' : 'ChevronDown'}
        styles={{ root: { fontSize: 12, color: theme.palette.neutralSecondary, flexShrink: 0 } }}
      />

      <div style={{ flex: 1, minWidth: 0 }}>
        <Text
          variant="smallPlus"
          styles={{ root: { display: 'block', fontWeight: 700, color: theme.semanticColors.bodyText } }}
        >
          {group.name}
        </Text>
        <Text
          variant="xSmall"
          styles={{
            root: {
              display: 'block',
              color: theme.palette.neutralSecondary,
              overflow: 'hidden',
              textOverflow: 'ellipsis',
              whiteSpace: 'nowrap'
            }
          }}
        >
          {group.key}
        </Text>
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: 6, flexShrink: 0 }}>
        <Text variant="xSmall" styles={{ root: { color: theme.palette.neutralSecondary } }}>
          {`${group.count} ${group.count === 1 ? 'page' : 'pages'}`}
        </Text>
        {staleCount > 0 && (
          <span style={pillStyle(theme.semanticColors.warningBackground, theme.semanticColors.warningText)}>
            {`${staleCount} Stale`}
          </span>
        )}
        {missingCount > 0 && (
          <span style={pillStyle(theme.semanticColors.errorBackground, theme.semanticColors.errorText)}>
            {`${missingCount} Missing`}
          </span>
        )}
        {abandonedCount > 0 && (
          <span style={pillStyle(theme.palette.neutralLighter, theme.palette.neutralSecondary)}>
            {`${abandonedCount} Abandoned`}
          </span>
        )}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Main table component
// ---------------------------------------------------------------------------

function DriftTable(props: IDriftTableProps): JSX.Element {
  const [sortState, setSortState] = React.useState<ISortState>(DEFAULT_SORT_STATE);

  // Hooks must be called before any early return.
  const sortedItems = React.useMemo(
    () => sortItemsGrouped(props.items, sortState),
    [props.items, sortState]
  );

  const groups = React.useMemo(() => buildGroups(sortedItems), [sortedItems]);

  const handleSort = (field: SortableField): void => {
    setSortState((current) => ({
      field,
      isDescending: current.field === field ? !current.isDescending : false
    }));
  };

  if (props.items.length === 0) {
    return <div>No pages match the current filter.</div>;
  }

  const groupProps: IGroupRenderProps = {
    onRenderHeader: (headerProps?: IGroupHeaderProps): JSX.Element | null => {
      if (!headerProps?.group) return null;
      return <DriftGroupHeader groupHeaderProps={headerProps} allItems={sortedItems} />;
    }
  };

  const columns: IColumn[] = [
    {
      key: 'title',
      name: 'Page',
      fieldName: 'Title',
      minWidth: 180,
      maxWidth: 280,
      isResizable: true,
      isSorted: sortState.field === 'Title',
      isSortedDescending: sortState.field === 'Title' && sortState.isDescending,
      onColumnClick: () => handleSort('Title'),
      onRender: (item: IDriftItem) =>
        item.DefaultPageUrl ? (
          <Link href={item.DefaultPageUrl} target="_blank">
            {item.Title}
          </Link>
        ) : (
          <span>{item.Title}</span>
        )
    },
    {
      key: 'status',
      name: 'Status',
      fieldName: 'DriftStatus',
      minWidth: 90,
      maxWidth: 110,
      isResizable: true,
      isSorted: sortState.field === 'DriftStatus',
      isSortedDescending: sortState.field === 'DriftStatus' && sortState.isDescending,
      onColumnClick: () => handleSort('DriftStatus'),
      onRender: (item: IDriftItem) => <StatusPill status={item.DriftStatus} />
    },
    {
      key: 'daysDrift',
      name: 'Days Drift',
      fieldName: 'DaysDrift',
      minWidth: 80,
      maxWidth: 100,
      isResizable: true,
      isSorted: sortState.field === 'DaysDrift',
      isSortedDescending: sortState.field === 'DaysDrift' && sortState.isDescending,
      onColumnClick: () => handleSort('DaysDrift')
    },
    {
      key: 'enModified',
      name: 'EN Modified',
      minWidth: 100,
      maxWidth: 120,
      onRender: (item: IDriftItem) => <span>{formatDate(item.DefaultPageModified)}</span>
    },
    {
      key: 'frModified',
      name: 'FR Modified',
      minWidth: 100,
      maxWidth: 120,
      onRender: (item: IDriftItem) => <span>{formatDate(item.TranslationModified)}</span>
    },
    {
      key: 'translator',
      name: 'Translator',
      fieldName: 'TranslatorName',
      minWidth: 100,
      maxWidth: 150,
      onRender: (item: IDriftItem) => <span>{item.TranslatorName || '—'}</span>
    },
    {
      key: 'frPage',
      name: 'FR Page',
      minWidth: 90,
      maxWidth: 110,
      onRender: (item: IDriftItem) =>
        item.TranslationPageUrl ? (
          <Link href={item.TranslationPageUrl} target="_blank">
            Open
          </Link>
        ) : (
          <span>—</span>
        )
    }
  ];

  return (
    <DetailsList
      items={sortedItems}
      columns={columns}
      selectionMode={SelectionMode.none}
      getKey={(item: IDriftItem) => `${item.PageGuid}|${item.TranslationLanguage}`}
      groups={groups}
      groupProps={groupProps}
    />
  );
}

export default DriftTable;
