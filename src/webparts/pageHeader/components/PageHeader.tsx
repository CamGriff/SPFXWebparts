import * as React from 'react';
import styles from './PageHeader.module.scss';
import type { IPageHeaderProps } from './IPageHeaderProps';
import Breadcrumb from './Breadcrumb';

const PageHeader: React.FC<IPageHeaderProps> = ({
  siteTitle,
  searchBoxPlaceholder,
  searchPageUrl,
  backgroundImageUrl,
  seasonalEnabled,
  seasonalLabel,
  seasonalUrl,
  context
}) => {

  const [query, setQuery] = React.useState<string>('');

  const siteCollectionUrl: string = context.pageContext.site.absoluteUrl;
  const serverRequestPath: string = context.pageContext.site.serverRequestPath;
  const isHomePage: boolean = context.pageContext.web.absoluteUrl === `${siteCollectionUrl}/SitePages/Home.aspx`
    || serverRequestPath === '/';

  const handleSearch = (): void => {
    if (!searchPageUrl || !query.trim()) return;
    window.location.href = `${searchPageUrl}?q=${encodeURIComponent(query)}`;
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === 'Enter') handleSearch();
  };

  return (
    <div
      className={styles.pageHeader}
      style={{ backgroundImage: backgroundImageUrl ? `url(${backgroundImageUrl})` : 'none' }}
    >
      <a href={siteCollectionUrl} className={styles.siteTitle}>
        {siteTitle}
      </a>

      {!isHomePage && (
        <Breadcrumb serverRequestPath={serverRequestPath} />
      )}

      <div className={styles.searchRow}>
        <div className={styles.searchContainer}>
          <input
            type="text"
            className={styles.searchInput}
            placeholder={searchBoxPlaceholder || ''}
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={handleKeyDown}
            aria-label={searchBoxPlaceholder || 'Search'}
          />
          <button
            className={styles.searchButton}
            onClick={handleSearch}
            aria-label="Search"
          >
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <circle cx="11" cy="11" r="8" />
              <line x1="21" y1="21" x2="16.65" y2="16.65" />
            </svg>
          </button>
        </div>

        {seasonalEnabled && seasonalLabel && seasonalUrl && (
          
          <a
            href={seasonalUrl}
            className={styles.seasonalButton}
            aria-label={seasonalLabel}
          >
            {seasonalLabel}
          </a>
        )}
      </div>
    </div>
  );
};

export default PageHeader;