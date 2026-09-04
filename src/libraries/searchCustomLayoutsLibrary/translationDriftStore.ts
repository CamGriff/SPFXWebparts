import { PageContext } from '@microsoft/sp-page-context';

export interface IDriftRecord {
  driftStatus: string;
  daysDrift: number;
}

// Keyed by normalized PageGuid (braces stripped, lowercased) so it matches
// regardless of whether the GUID arrives bare (as PageGuid is stored in the
// TranslationDrift list) or wrapped in braces (as UniqueID comes back from
// search results, e.g. "{39dd4f35-...}").
let driftMap: Map<string, IDriftRecord> = new Map();
let loaded = false;
let loadPromise: Promise<void> | undefined;

function normalizeGuid(guid: string): string {
  return (guid || '').replace(/[{}]/g, '').toLowerCase();
}

function whenScopeFinished(serviceScope: any): Promise<void> {
  return new Promise<void>((resolve) => {
    if (!serviceScope || typeof serviceScope.whenFinished !== 'function') {
      resolve();
      return;
    }
    serviceScope.whenFinished(() => resolve());
  });
}

/**
 * Fetches the TranslationDrift list once per page load and builds a
 * GUID-keyed lookup map.
 *
 * Deliberately bypasses @pnp/sp's SPFx() context helper, which repeatedly
 * failed with "Cannot read properties of undefined (reading 'web')" when
 * given a bare BaseLayout serviceScope. Uses the same plain fetch() +
 * credentials:"same-origin" pattern already proven working elsewhere in
 * this codebase (see ActionWebComponent/CustomComponent.tsx's REST calls),
 * authenticating via the page's own session cookie rather than PnP JS's
 * own context resolution.
 */
export async function loadTranslationDrift(serviceScope: any, listName: string): Promise<void> {
  if (loaded) {
    return;
  }
  if (loadPromise) {
    return loadPromise;
  }

  loadPromise = (async () => {
    try {
      await whenScopeFinished(serviceScope);

      let webUrl: string | undefined;
      try {
        const pageContext: PageContext = serviceScope.consume(PageContext.serviceKey);
        webUrl = pageContext?.web?.absoluteUrl;
      } catch (consumeError) {
        console.error('[SearchCustomLayoutsLibrary] Could not consume PageContext from serviceScope:', consumeError);
      }

      if (!webUrl) {
        console.error('[SearchCustomLayoutsLibrary] No web URL available — aborting TranslationDrift fetch.');
        return;
      }

      const encodedListName = encodeURIComponent(listName);
      const url = `${webUrl}/_api/web/lists/getbytitle('${encodedListName}')/items`
        + `?$select=PageGuid,DriftStatus,DaysDrift&$top=5000`;

      const response = await fetch(url, {
        method: 'GET',
        headers: { accept: 'application/json;odata=nometadata' },
        credentials: 'same-origin'
      });

      if (!response.ok) {
        console.error(`[SearchCustomLayoutsLibrary] TranslationDrift fetch failed: ${response.status} ${response.statusText}`);
        return;
      }

      const json = await response.json();
      const items: any[] = json?.value || [];

      const map = new Map<string, IDriftRecord>();
      for (const item of items) {
        if (!item.PageGuid) {
          continue;
        }
        map.set(normalizeGuid(item.PageGuid), {
          driftStatus: item.DriftStatus,
          daysDrift: item.DaysDrift
        });
      }

      driftMap = map;
      loaded = true;
      console.log('[SearchCustomLayoutsLibrary] TranslationDrift loaded, records:', map.size);
    } catch (error) {
      console.error('[SearchCustomLayoutsLibrary] Failed to load TranslationDrift list:', error);
    }
  })();

  return loadPromise;
}

/**
 * Synchronous lookup for use inside Handlebars helpers. Returns undefined
 * if the page has no drift record (e.g. a page not yet tracked by a scan).
 */
export function getDriftRecord(pageGuid: string): IDriftRecord | undefined {
  return driftMap.get(normalizeGuid(pageGuid));
}