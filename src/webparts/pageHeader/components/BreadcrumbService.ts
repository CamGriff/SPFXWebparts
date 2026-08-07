import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/items";

export interface IBreadcrumbTranslation {
  Title: string; // internal name of the "FolderName" column — the URL path segment
  FolderPathFR?: string;
  [key: string]: string | undefined;
}

const CACHE_PREFIX = 'bc_trans_';

// Bump this any time the fields read from BreadcrumbTranslations change —
// i.e. a field is added to, removed from, or renamed in IBreadcrumbTranslation
// or the $select clause below. This changes the cache key, so old sessionStorage
// entries with the previous shape are never read again instead of silently
// serving stale/incomplete data.
const TRANSLATIONS_SCHEMA_VERSION = 2;

export async function loadTranslations(context: WebPartContext): Promise<IBreadcrumbTranslation[]> {
  const siteUrl = context.pageContext.site.absoluteUrl;
  const cacheKey = `${CACHE_PREFIX}v${TRANSLATIONS_SCHEMA_VERSION}_${siteUrl}`;

  try {
    const cached = sessionStorage.getItem(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch { /* ignore corrupt cache */ }

  try {
    const url = `${siteUrl}/_api/web/lists/getByTitle('BreadcrumbTranslations')/items?$select=Title,LabelEN,LabelFR,FolderPath,FolderPathFR&$top=5000`;
    const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) return [];

    const data = await response.json();
    const items: IBreadcrumbTranslation[] = (data.value ?? []).map((item: Record<string, string>) => {
      const entry: IBreadcrumbTranslation = {
        Title: item.Title || '',
        FolderPath: item.FolderPath || '',
        FolderPathFR: item.FolderPathFR || ''
      };
      for (const key of Object.keys(item)) {
        if (key.startsWith('Label')) entry[key] = item[key];
      }
      return entry;
    });

    sessionStorage.setItem(cacheKey, JSON.stringify(items));
    return items;
  } catch {
    return [];
  }
}

export function getTranslationItem(
  translations: IBreadcrumbTranslation[],
  slug: string
): IBreadcrumbTranslation | undefined {
  const normalized = decodeURIComponent(slug).toLowerCase();
  return translations.find(t => (t.Title || '').toLowerCase() === normalized);
}

export function translateSlug(
  translations: IBreadcrumbTranslation[],
  slug: string,
  langKey: string
): string {
  const normalized = decodeURIComponent(slug).toLowerCase();
  const item = translations.find(t => (t.Title || '').toLowerCase() === normalized);
  if (!item) return decodeURIComponent(slug);
  return item[`Label${langKey.toUpperCase()}`] || item['LabelEN'] || decodeURIComponent(slug);
}

export async function getPageTitle(context: WebPartContext, serverRequestPath: string): Promise<string> {
  try {
    const sp = spfi().using(SPFx(context));
    const item = await sp.web
      .getFileByServerRelativePath(serverRequestPath)
      .getItem<{ Title: string }>("Title");
    return item.Title;
  } catch {
    return '';
  }
}
