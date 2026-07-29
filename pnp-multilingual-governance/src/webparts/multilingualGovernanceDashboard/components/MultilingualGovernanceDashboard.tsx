


import * as React from 'react';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { SPFI } from '@pnp/sp';
import { Stack, PrimaryButton, DefaultButton, Toggle, Spinner, SpinnerSize, MessageBar, MessageBarType } from '@fluentui/react';

import { IMultilingualGovernanceDashboardProps } from './IMultilingualGovernanceDashboardProps';
import { IDriftItem } from '../../../models/IDriftItem';
import { IScanWarning } from '../../../models/IScanWarning';
import { runDriftScan } from '../../../services/DriftService';
import SummaryTiles, { SummaryTileKey } from './dashboard/SummaryTiles';
import DriftTable from './dashboard/DriftTable';
import TranslatorPanel from './dashboard/TranslatorPanel';
import ConfigPanel from './config/ConfigPanel';

const TRANSLATION_DRIFT_LIST_TITLE = 'TranslationDrift';
const MAX_ITEMS_PER_REQUEST = 5000;

function csvEscape(value: string): string {
  if (value.includes(',') || value.includes('"') || value.includes('\n')) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}

function csvDate(value: string | undefined): string {
  if (!value) return '';
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? '' : parsed.toLocaleDateString();
}

function exportToCsv(items: IDriftItem[]): void {
  const header = ['Page', 'Status', 'Days Drift', 'EN Modified', 'FR Modified', 'Translator', 'FR Page URL'];
  const rows = items.map((item) => [
    csvEscape(item.Title),
    csvEscape(item.DriftStatus),
    String(item.DaysDrift),
    csvDate(item.DefaultPageModified),
    csvDate(item.TranslationModified),
    csvEscape(item.TranslatorName ?? ''),
    csvEscape(item.TranslationPageUrl)
  ]);
  const csv = [header.join(','), ...rows.map((r) => r.join(','))].join('\r\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = 'translation-drift.csv';
  anchor.click();
  URL.revokeObjectURL(url);
}

interface IRawDriftItem extends Omit<IDriftItem, 'DefaultPageUrl' | 'TranslationPageUrl'> {
  DefaultPageUrl: { Url: string } | undefined;
  TranslationPageUrl: { Url: string } | undefined;
}

async function loadDriftItems(sp: SPFI, showInSync: boolean): Promise<IDriftItem[]> {
  let query = sp.web.lists
    .getByTitle(TRANSLATION_DRIFT_LIST_TITLE)
    .items.select(
      'Title',
      'DefaultPageUrl',
      'DefaultPageModified',
      'TranslationLanguage',
      'TranslationPageUrl',
      'TranslationModified',
      'DaysDrift',
      'DriftStatus',
      'TranslatorName',
      'TranslatorEmail',
      'SiteUrl',
      'PageGuid'
    )
    .orderBy('DaysDrift', false)
    .top(MAX_ITEMS_PER_REQUEST);

  // In Sync pages need no action, so exclude them from the default payload.
  if (!showInSync) {
    query = query.filter(`DriftStatus ne '${DriftStatus.InSync}'`);
  }

  const raw = (await query()) as IRawDriftItem[];

  return raw.map((item) => ({
    ...item,
    DefaultPageUrl: item.DefaultPageUrl?.Url ?? '',
    TranslationPageUrl: item.TranslationPageUrl?.Url ?? ''
  }));
}

function MultilingualGovernanceDashboard(props: IMultilingualGovernanceDashboardProps): JSX.Element {
  const [items, setItems] = React.useState<IDriftItem[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [isScanning, setIsScanning] = React.useState<boolean>(false);
  const [scanError, setScanError] = React.useState<string | undefined>(undefined);
  const [scanWarnings, setScanWarnings] = React.useState<IScanWarning[]>([]);
  const [loadError, setLoadError] = React.useState<string | undefined>(undefined);
  const [activeFilter, setActiveFilter] = React.useState<SummaryTileKey>('All');
  const [showConfigPanel, setShowConfigPanel] = React.useState<boolean>(false);

  const loadItems = React.useCallback(async (): Promise<void> => {
    console.log('[MGov] loadItems: start');
    setIsLoading(true);
    setLoadError(undefined);

    try {
      const driftItems = await loadDriftItems(props.sp);
      console.log('[MGov] loadItems: succeeded, count', driftItems.length);
      setItems(driftItems);
    } catch (error) {
      console.error('[MGov] loadItems: failed — setting loadError', error);
      setLoadError('Failed to load translation drift data.');
    } finally {
      setIsLoading(false);
    }
  }, [props.sp]);

  React.useEffect(() => {
    void loadItems();
  }, [loadItems]);

  const handleRunScan = async (): Promise<void> => {
    console.log('[MGov] handleRunScan: start — clearing scanError');
    setIsScanning(true);
    setScanError(undefined);
    setScanWarnings([]);

    try {
      console.log('[MGov] handleRunScan: calling runDriftScan');
      const result = await runDriftScan(props.sp);
      console.log('[MGov] handleRunScan: runDriftScan succeeded, warnings', result.warnings.length);
      setScanWarnings(result.warnings);
    } catch (error) {
      console.error('[MGov] handleRunScan: catch block — setting scanError', error);
      setScanError('Drift scan failed. Please try again.');
      setIsScanning(false);
      return;
    }

    console.log('[MGov] handleRunScan: scan succeeded, calling loadItems');
    setIsScanning(false);
    await loadItems();
    console.log('[MGov] handleRunScan: complete');
  };

  const filteredItems = activeFilter === 'All' ? items : items.filter((item) => item.DriftStatus === activeFilter);

  return (
    <Stack tokens={{ childrenGap: 16 }} styles={{ root: { padding: 16 } }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" wrap tokens={{ childrenGap: 12 }}>
        <h2 style={{ margin: 0 }}>Translation Drift Dashboard</h2>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text="Run Scan"
            onClick={() => {
              void handleRunScan();
            }}
            disabled={isScanning || isLoading}
          />
          <DefaultButton
            text="Refresh"
            onClick={() => {
              void loadItems();
            }}
            disabled={isScanning || isLoading}
          />
          <DefaultButton
            text="Export"
            onClick={() => exportToCsv(filteredItems)}
            disabled={isScanning || isLoading || filteredItems.length === 0}
          />
          <Toggle
            label="Manage sites"
            inlineLabel
            checked={showConfigPanel}
            onChange={(_event, checked) => setShowConfigPanel(!!checked)}
          />
        </Stack>
      </Stack>

      {scanError && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setScanError(undefined)}>
          {scanError}
        </MessageBar>
      )}
      {loadError && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setLoadError(undefined)}>
          {loadError}
        </MessageBar>
      )}
      {scanWarnings.map((warning) => (
        <MessageBar
          key={warning.siteUrl}
          messageBarType={MessageBarType.warning}
          onDismiss={() => setScanWarnings((current) => current.filter((w) => w.siteUrl !== warning.siteUrl))}
        >
          {`Could not scan the following site: ${warning.siteUrl} — ${warning.error}`}
        </MessageBar>
      ))}

      {isScanning && <Spinner size={SpinnerSize.medium} label="Scanning sites for translation drift..." />}

      {showConfigPanel && <ConfigPanel sp={props.sp} />}

      {isLoading ? (
        <Spinner size={SpinnerSize.large} label="Loading translation drift data..." />
      ) : (
        <>
          <SummaryTiles items={items} activeFilter={activeFilter} onFilterChange={setActiveFilter} />
          <TranslatorPanel allItems={items} sp={props.sp} graphClient={props.graphClient} />
          <DriftTable items={filteredItems} />
        </>
      )}
    </Stack>
  );
}

export default MultilingualGovernanceDashboard;
