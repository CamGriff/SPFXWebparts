import * as React from 'react';
import styles from './PageHeader.module.scss';

interface IBreadcrumbProps {
  serverRequestPath: string;
  siteCollectionUrl: string;
}

const Breadcrumb: React.FC<IBreadcrumbProps> = ({ serverRequestPath, siteCollectionUrl }) => {

  const segments: { label: string; url: string }[] = React.useMemo(() => {
    // Extract the path after SitePages
    const sitePagesMarker = '/SitePages/';
    const sitePagesIndex = serverRequestPath.indexOf(sitePagesMarker);
    if (sitePagesIndex === -1) return [];

    const relativePath = serverRequestPath.slice(sitePagesIndex + sitePagesMarker.length);
    const parts = relativePath.split('/');

    // Strip the last part (the page filename itself)
    const folders = parts.slice(0, -1);

    // Build cumulative URLs for each folder step
    return folders.map((folder, index) => {
      const url = `${siteCollectionUrl}/SitePages/${folders.slice(0, index + 1).join('/')}`; 
      return { label: decodeURIComponent(folder), url };
    });
  }, [serverRequestPath, siteCollectionUrl]);

  if (segments.length === 0) return null;

  return (
    <nav className={styles.breadcrumb} aria-label="Breadcrumb">
      {segments.map((segment, index) => (
        <React.Fragment key={index}>
          {index > 0 && <span className={styles.breadcrumbSeparator}>/</span>}
          <a href={segment.url} className={styles.breadcrumbLink}>
            {segment.label}
          </a>
        </React.Fragment>
      ))}
    </nav>
  );
};

export default Breadcrumb;