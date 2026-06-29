import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { IWeb, Web } from '@pnp/sp/webs';

import { DriftStatus } from '../models/DriftStatus';
import { IConfigItem } from '../models/IConfigItem';
import { IDriftItem } from '../models/IDriftItem';
import { IScanWarning } from '../models/IScanWarning';
import { getConfigs } from './ConfigService';

export interface IDriftScanResult {
  items: IDriftItem[];
  warnings: IScanWarning[];
}

const TRANSLATION_DRIFT_LIST_TITLE = 'TranslationDrift';
const SITE_PAGES_LIBRARY_TITLE = 'Site Pages';
const APPROVED_MODERATION_STATUS = 0;
const EXCLUDED_PATH_SEGMENT = '/fr/';
const MS_PER_DAY = 24 * 60 * 60 * 1000;
const MAX_ITEMS_PER_REQUEST = 5000;

interface ISitePageData {
  Id: number;
  Title: string;
  FileRef: string;
  UniqueId: string;
  Modified: string;
  OData__SPIsTranslation: boolean | undefined;
  OData__SPTranslationLanguage: string | undefined;
  OData__SPTranslationSourceItemId: string | undefined;
  OData__ModerationStatus: number;
}

interface IExistingDriftRow {
  Id: number;
  PageGuid: string;
  TranslationLanguage: string;
}

export async function runDriftScan(sp: SPFI): Promise<IDriftScanResult> {
  const configs = await getConfigs(sp);
  console.log('[MGov] runDriftScan: active configs', configs.length);
  const driftItems: IDriftItem[] = [];
  const warnings: IScanWarning[] = [];

  for (const config of configs) {
    try {
      console.log('[MGov] runDriftScan: scanning', config.SiteUrl);
      const siteDriftItems = await scanSiteForLanguage(sp, config);
      console.log('[MGov] runDriftScan: site ok, items', siteDriftItems.length);
      driftItems.push(...siteDriftItems);
    } catch (error) {
      console.error(`[MGov] runDriftScan: site failed ${config.SiteUrl}`, error);
      warnings.push({ siteUrl: config.SiteUrl, error: 'Access denied or site unavailable' });
    }
  }

  console.log('[MGov] runDriftScan: upsert start, total items', driftItems.length);
  await upsertDriftItems(sp, driftItems);
  console.log('[MGov] runDriftScan: upsert complete, warnings', warnings.length);

  return { items: driftItems, warnings };
}

async function scanSiteForLanguage(sp: SPFI, config: IConfigItem): Promise<IDriftItem[]> {
  const web = Web([sp.web, config.SiteUrl]);
  const pages = await getSitePages(web);
  const origin = new URL(config.SiteUrl).origin;

  const sourcePages = pages.filter((page) => isSourcePage(page));
  const translationPages = pages.filter((page) => isTranslationPage(page, config.Language));

  const translationsBySourceId = new Map<string, ISitePageData>();
  for (const translation of translationPages) {
    if (translation.OData__SPTranslationSourceItemId) {
      translationsBySourceId.set(translation.OData__SPTranslationSourceItemId.toLowerCase(), translation);
    }
  }

  const matchedTranslationIds = new Set<number>();
  const driftItems: IDriftItem[] = [];

  for (const sourcePage of sourcePages) {
    const translation = translationsBySourceId.get(sourcePage.UniqueId.toLowerCase());
    if (translation) {
      matchedTranslationIds.add(translation.Id);
    }
    driftItems.push(buildSourceDriftItem(sourcePage, translation, config, origin));
  }

  for (const translation of translationPages) {
    if (!matchedTranslationIds.has(translation.Id)) {
      driftItems.push(buildOrphanedDriftItem(translation, config, origin));
    }
  }

  return driftItems;
}

async function getSitePages(web: IWeb): Promise<ISitePageData[]> {
  const pages: ISitePageData[] = await web.lists
    .getByTitle(SITE_PAGES_LIBRARY_TITLE)
    .items.select(
      'Id',
      'Title',
      'FileRef',
      'UniqueId',
      'Modified',
      'OData__SPIsTranslation',
      'OData__SPTranslationLanguage',
      'OData__SPTranslationSourceItemId',
      'OData__ModerationStatus'
    )
    .top(MAX_ITEMS_PER_REQUEST)();

  return pages;
}

function isSourcePage(page: ISitePageData): boolean {
  return (
    page.FileRef.toLowerCase().endsWith('.aspx') &&
    page.OData__SPIsTranslation !== true &&
    !page.FileRef.toLowerCase().includes(EXCLUDED_PATH_SEGMENT) &&
    !!page.Title &&
    page.OData__ModerationStatus === APPROVED_MODERATION_STATUS
  );
}

function isTranslationPage(page: ISitePageData, language: string): boolean {
  return page.OData__SPIsTranslation === true && (page.OData__SPTranslationLanguage || '').toLowerCase() === language.toLowerCase();
}

function buildSourceDriftItem(
  sourcePage: ISitePageData,
  translation: ISitePageData | undefined,
  config: IConfigItem,
  origin: string
): IDriftItem {
  const base = {
    Title: sourcePage.Title,
    DefaultPageTitle: sourcePage.Title,
    DefaultPageUrl: combineUrl(origin, sourcePage.FileRef),
    DefaultPageModified: sourcePage.Modified,
    TranslationLanguage: config.Language,
    TranslatorName: config.TranslatorName,
    TranslatorEmail: config.TranslatorEmail,
    SiteUrl: config.SiteUrl,
    LastChecked: new Date().toISOString(),
    PageGuid: sourcePage.UniqueId,
    NudgeSent: false,
    NudgeDate: undefined
  };

  if (!translation) {
    return {
      ...base,
      TranslationPageUrl: '',
      TranslationModified: undefined,
      DaysDrift: 0,
      DriftStatus: DriftStatus.Missing
    };
  }

  if (translation.OData__ModerationStatus !== APPROVED_MODERATION_STATUS) {
    return {
      ...base,
      TranslationPageUrl: combineUrl(origin, translation.FileRef),
      TranslationModified: translation.Modified,
      DaysDrift: 0,
      DriftStatus: DriftStatus.Abandoned
    };
  }

  const daysDrift = calculateDaysDrift(sourcePage.Modified, translation.Modified);
  const status = daysDrift > config.StaleDays ? DriftStatus.Stale : DriftStatus.InSync;

  return {
    ...base,
    TranslationPageUrl: combineUrl(origin, translation.FileRef),
    TranslationModified: translation.Modified,
    DaysDrift: daysDrift,
    DriftStatus: status
  };
}

function buildOrphanedDriftItem(translation: ISitePageData, config: IConfigItem, origin: string): IDriftItem {
  return {
    Title: translation.Title,
    DefaultPageTitle: '',
    DefaultPageUrl: '',
    DefaultPageModified: '',
    TranslationLanguage: config.Language,
    TranslationPageUrl: combineUrl(origin, translation.FileRef),
    TranslationModified: translation.Modified,
    DaysDrift: 0,
    DriftStatus: DriftStatus.Orphaned,
    TranslatorName: config.TranslatorName,
    TranslatorEmail: config.TranslatorEmail,
    SiteUrl: config.SiteUrl,
    LastChecked: new Date().toISOString(),
    PageGuid: translation.UniqueId,
    NudgeSent: false,
    NudgeDate: undefined
  };
}

function calculateDaysDrift(defaultModified: string, translationModified: string): number {
  const rawDaysDrift = Math.floor((new Date(defaultModified).getTime() - new Date(translationModified).getTime()) / MS_PER_DAY);
  return Math.max(rawDaysDrift, 0);
}

function combineUrl(origin: string, serverRelativeUrl: string): string {
  return `${origin}${serverRelativeUrl}`;
}

async function upsertDriftItems(sp: SPFI, items: IDriftItem[]): Promise<void> {
  if (items.length === 0) {
    return;
  }

  const list = sp.web.lists.getByTitle(TRANSLATION_DRIFT_LIST_TITLE);
  const existingRows: IExistingDriftRow[] = await list.items
    .select('Id', 'PageGuid', 'TranslationLanguage')
    .top(MAX_ITEMS_PER_REQUEST)();

  console.log(`[MGov] upsertDriftItems: loaded ${existingRows.length} existing rows`);
  const existingKeys = existingRows.map((r) => getDriftRowKey(r.PageGuid, r.TranslationLanguage));
  console.log('[MGov] upsertDriftItems: existing keys (first 5):', existingKeys.slice(0, 5));

  const existingIdsByKey = new Map<string, number>();
  for (const row of existingRows) {
    const key = getDriftRowKey(row.PageGuid, row.TranslationLanguage);
    existingIdsByKey.set(key, row.Id);
    if (row.PageGuid !== row.PageGuid.trim() || row.TranslationLanguage !== row.TranslationLanguage.trim()) {
      console.warn('[MGov] upsertDriftItems: whitespace detected in existing row', JSON.stringify(row.PageGuid), JSON.stringify(row.TranslationLanguage));
    }
  }

  for (const item of items) {
    const incomingKey = getDriftRowKey(item.PageGuid, item.TranslationLanguage);
    const existingId = existingIdsByKey.get(incomingKey);

    console.log(
      `[MGov] upsert: key="${incomingKey}" rawGuid=${JSON.stringify(item.PageGuid)} rawLang=${JSON.stringify(item.TranslationLanguage)} → ${existingId !== undefined ? `UPDATE id=${existingId}` : 'CREATE'}`
    );

    try {
      if (existingId !== undefined) {
        await list.items.getById(existingId).update(toUpdateFields(item));
      } else {
        await list.items.add(toCreateFields(item));
      }
    } catch (error) {
      console.error('[MGov] upsertDriftItems: failed on item', item.PageGuid, item.DriftStatus, 'DefaultPageModified:', JSON.stringify(item.DefaultPageModified), error);
      throw error;
    }
  }
}

function getDriftRowKey(pageGuid: string, translationLanguage: string): string {
  return `${pageGuid.toLowerCase()}|${translationLanguage.toLowerCase()}`;
}

function urlField(url: string | undefined): { Url: string; Description: string } | undefined {
  if (!url) return undefined;
  return { Url: url, Description: url };
}

// Converts empty strings to undefined so JSON.stringify omits the key entirely.
// SharePoint rejects "" for DateTime columns — undefined is safely omitted.
function dateField(value: string | undefined): string | undefined {
  return value || undefined;
}

function toCreateFields(item: IDriftItem): Record<string, unknown> {
  return {
    ...item,
    DefaultPageUrl: urlField(item.DefaultPageUrl),
    TranslationPageUrl: urlField(item.TranslationPageUrl),
    DefaultPageModified: dateField(item.DefaultPageModified),
    TranslationModified: dateField(item.TranslationModified),
    NudgeDate: dateField(item.NudgeDate)
  };
}

function toUpdateFields(item: IDriftItem): Record<string, unknown> {
  return {
    Title: item.Title,
    DefaultPageTitle: item.DefaultPageTitle,
    DefaultPageUrl: urlField(item.DefaultPageUrl),
    DefaultPageModified: dateField(item.DefaultPageModified),
    TranslationLanguage: item.TranslationLanguage,
    TranslationPageUrl: urlField(item.TranslationPageUrl),
    TranslationModified: dateField(item.TranslationModified),
    DaysDrift: item.DaysDrift,
    DriftStatus: item.DriftStatus,
    TranslatorName: item.TranslatorName,
    TranslatorEmail: item.TranslatorEmail,
    SiteUrl: item.SiteUrl,
    LastChecked: item.LastChecked,
    PageGuid: item.PageGuid
  };
}
