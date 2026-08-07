import * as React from 'react';
import styles from './PageHeader.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { loadTranslations, translateSlug, IBreadcrumbTranslation, getTranslationItem, getPageTitle } from './BreadcrumbService';

interface IBreadcrumbProps {
  serverRequestPath: string;
  context: WebPartContext;
  siteCollectionUrl: string;
}

interface IPathParts {
  parentFolderRaw: string;
  pageRaw: string;
  parentFolderUrl: string;
}

function parsePath(
  serverRequestPath: string,
  siteCollectionUrl: string,
  translations: IBreadcrumbTranslation[]
): IPathParts | null {
  const marker = '/SitePages/';
  const idx = serverRequestPath.indexOf(marker);
  if (idx === -1) return null;

  const relative = serverRequestPath.slice(idx + marker.length);
  let parts = relative.split('/').filter(Boolean);
  if (parts.length < 2) return null;

  // The locale segment, when present, is the folder immediately before the
  // filename (e.g. SitePages/IT/fr/Page.aspx), not a fixed site-root prefix.
  const localePattern = /^[a-z]{2}(-[a-z]{2})?$/i;
  const lastFolderIndex = parts.length - 2;
  if (localePattern.test(parts[lastFolderIndex]) && !getTranslationItem(translations, parts[lastFolderIndex])) {
    parts = [...parts.slice(0, lastFolderIndex), parts[parts.length - 1]];
  }

  if (parts.length < 2) return null;

  const pageRaw = parts[parts.length - 1].replace(/\.aspx$/i, '');
  const folderParts = parts.slice(0, -1);
  const parentFolderRaw = folderParts[folderParts.length - 1];
  const parentFolderUrl = `${siteCollectionUrl}/SitePages/${folderParts.join('/')}`;

  return { parentFolderRaw, pageRaw, parentFolderUrl };
}

const Breadcrumb: React.FC<IBreadcrumbProps> = ({ serverRequestPath, context, siteCollectionUrl }) => {
  const [translations, setTranslations] = React.useState<IBreadcrumbTranslation[]>([]);
  const [translationsLoaded, setTranslationsLoaded] = React.useState<boolean>(false);
  const [pageLabel, setPageLabel] = React.useState<string>('');

  React.useEffect(() => {
    loadTranslations(context)
      .then(items => {
        setTranslations(items);
        setTranslationsLoaded(true);
      })
      .catch(() => setTranslationsLoaded(true));
  }, [context]);

  const parts = React.useMemo(
    () => (translationsLoaded ? parsePath(serverRequestPath, siteCollectionUrl, translations) : null),
    [serverRequestPath, siteCollectionUrl, translations, translationsLoaded]
  );

  React.useEffect(() => {
    if (!parts) return;
    getPageTitle(context, serverRequestPath)
      .then(title => setPageLabel(title || parts.pageRaw))
      .catch(() => setPageLabel(parts.pageRaw));
  }, [context, serverRequestPath, parts]);

  if (!parts) return null;

  const langKey = context.pageContext.cultureInfo.currentUICultureName.split('-')[0];
  const folderItem = getTranslationItem(translations, parts.parentFolderRaw);
  const folderLabel = translateSlug(translations, parts.parentFolderRaw, langKey);
  const folderUrl =
    (langKey === 'fr' ? folderItem?.FolderPathFR : folderItem?.FolderPath) ||
    folderItem?.FolderPath ||
    parts.parentFolderUrl;

  return (
    <nav className={styles.breadcrumb} aria-label="Breadcrumb">
      <a href={folderUrl} className={styles.breadcrumbLink}>
        {folderLabel}
      </a>
      <span className={styles.breadcrumbSeparator}>/</span>
      <span className={styles.breadcrumbCurrent}>{pageLabel}</span>
    </nav>
  );
};

export default Breadcrumb;
