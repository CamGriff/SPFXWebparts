import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

export interface IBreadcrumbTranslation {
  Title: string; // internal name of the "FolderName" column — the URL path segment
  [key: string]: string;
}

const CACHE_PREFIX = 'bc_trans_';

export async function loadTranslations(context: WebPartContext): Promise<IBreadcrumbTranslation[]> {
  const siteUrl = context.pageContext.site.absoluteUrl;
  const cacheKey = `${CACHE_PREFIX}${siteUrl}`;

  try {
    const cached = sessionStorage.getItem(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch { /* ignore corrupt cache */ }

  try {
    const url = `${siteUrl}/_api/web/lists/getByTitle('BreadcrumbTranslations')/items?$select=Title,LabelEN,LabelFR,FolderPath&$top=5000`;
    const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) return [];

    const data = await response.json();
    const items: IBreadcrumbTranslation[] = (data.value ?? []).map((item: Record<string, string>) => {
      const entry: IBreadcrumbTranslation = { Title: item.Title || '', FolderPath: item.FolderPath || '' };
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
