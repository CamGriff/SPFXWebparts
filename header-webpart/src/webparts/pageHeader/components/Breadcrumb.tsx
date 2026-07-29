import * as React from 'react';
import styles from './PageHeader.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { loadTranslations, translateSlug, IBreadcrumbTranslation, getTranslationItem } from './BreadcrumbService';

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

function parsePath(serverRequestPath: string, siteCollectionUrl: string): IPathParts | null {
  const marker = '/SitePages/';
  const idx = serverRequestPath.indexOf(marker);
  if (idx === -1) return null;

  const relative = serverRequestPath.slice(idx + marker.length);
  const parts = relative.split('/').filter(Boolean);
  if (parts.length < 2) return null;

  const pageRaw = parts[parts.length - 1].replace(/\.aspx$/i, '');
  const folderParts = parts.slice(0, -1);
  const parentFolderRaw = folderParts[folderParts.length - 1];
  const parentFolderUrl = `${siteCollectionUrl}/SitePages/${folderParts.join('/')}`;

  return { parentFolderRaw, pageRaw, parentFolderUrl };
}

const Breadcrumb: React.FC<IBreadcrumbProps> = ({ serverRequestPath, context, siteCollectionUrl }) => {
  const [translations, setTranslations] = React.useState<IBreadcrumbTranslation[]>([]);

  React.useEffect(() => {
    loadTranslations(context).then(setTranslations).catch(() => undefined);
  }, [context]);

  const parts = React.useMemo(
    () => parsePath(serverRequestPath, siteCollectionUrl),
    [serverRequestPath, siteCollectionUrl]
  );

  if (!parts) return null;

  const langKey = context.pageContext.cultureInfo.currentUICultureName.split('-')[0];
  const folderItem = getTranslationItem(translations, parts.parentFolderRaw);
  const folderLabel = translateSlug(translations, parts.parentFolderRaw, langKey);
  const folderUrl = folderItem?.FolderPath || parts.parentFolderUrl;
  const pageLabel = translateSlug(translations, parts.pageRaw, langKey);

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
