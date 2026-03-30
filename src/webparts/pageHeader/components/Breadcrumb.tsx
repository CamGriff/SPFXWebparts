import * as React from 'react';
import styles from './PageHeader.module.scss';

interface IBreadcrumbProps {
  serverRequestPath: string;
}

const Breadcrumb: React.FC<IBreadcrumbProps> = ({ serverRequestPath }) => {

  const segments: string[] = React.useMemo(() => {
    const sitePagesMarker = '/SitePages/';
    const sitePagesIndex = serverRequestPath.indexOf(sitePagesMarker);
    if (sitePagesIndex === -1) return [];

    const relativePath = serverRequestPath.slice(sitePagesIndex + sitePagesMarker.length);
    const parts = relativePath.split('/');
    return parts.slice(0, -1).map(folder => decodeURIComponent(folder));
  }, [serverRequestPath]);

  if (segments.length === 0) return null;

  return (
    <nav className={styles.breadcrumb} aria-label="Breadcrumb">
      {segments.map((segment, index) => (
        <React.Fragment key={index}>
          {index > 0 && <span className={styles.breadcrumbSeparator}>/</span>}
          <span className={styles.breadcrumbLink}>{segment}</span>
        </React.Fragment>
      ))}
    </nav>
  );
};

export default Breadcrumb;