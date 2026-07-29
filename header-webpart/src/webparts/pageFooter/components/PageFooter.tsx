import * as React from 'react';
import styles from './PageFooter.module.scss';
import type { IPageFooterProps, IFooterLink } from './IPageFooterProps';

const PageFooter: React.FC<IPageFooterProps> = ({
  footerListName,
  linkedInUrl,
  youTubeUrl,
  logoUrl,
  context
}) => {

  const [links, setLinks] = React.useState<IFooterLink[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string>('');

  React.useEffect(() => {
    const siteUrl = context.pageContext.web.absoluteUrl;
    const apiUrl = `${siteUrl}/_api/web/lists/getbytitle('${footerListName}')/items?$select=Title,URL,Category,SortOrder&$orderby=Category,SortOrder`;

    fetch(apiUrl, {
      headers: {
        'Accept': 'application/json;odata=nometadata'
      },
      credentials: 'same-origin'
    })
      .then(res => {
        if (!res.ok) throw new Error(`Failed to fetch footer links: ${res.status}`);
        return res.json();
      })
      .then(data => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const mapped: IFooterLink[] = data.value.map((item: any) => ({
          title: item.Title,
          url: item.URL?.Url || item.URL || '',
          category: item.Category?.trim() || '',
          sortOrder: item.SortOrder
        }));
        setLinks(mapped);
        setLoading(false);
      })
      .catch(err => {
        setError(err.message);
        setLoading(false);
      });
  }, [footerListName]);

  // Group links by category
  const grouped = React.useMemo(() => {
    return links.reduce<Record<string, IFooterLink[]>>((acc, link) => {
      const cat = link.category || 'General';
      if (!acc[cat]) acc[cat] = [];
      acc[cat].push(link);
      return acc;
    }, {});
  }, [links]);

  const currentYear = new Date().getFullYear();

  const groupedEntries = (Object.entries(grouped) as [string, IFooterLink[]][]).slice(0, 4);
  const columnCount = groupedEntries.length;
  const spacerClass = columnCount >= 4 ? styles.spacerThird : styles.spacerHalf;

  if (loading) return <div className={styles.pageFooter} />;
  if (error) return <div className={styles.pageFooter}><p>{error}</p></div>;

  return (
    <footer className={styles.pageFooter}>
      <div className={styles.footerInner}>

        {logoUrl && (
          <img src={logoUrl} className={styles.logo} alt="Company logo" />
        )}

        <div className={spacerClass} />

        <div className={styles.linkColumns}>
  {groupedEntries.map(([category, categoryLinks]) => (
    <div key={category} className={styles.linkColumn}>
      <h4 className={styles.categoryTitle}>{category}</h4>
      <ul className={styles.linkList}>
        {categoryLinks.map((link: IFooterLink, index: number) => (
          <li key={index}>
            <a href={link.url} className={styles.footerLink}>
              {link.title}
            </a>
          </li>
        ))}
      </ul>
    </div>
  ))}
</div>

        <div className={styles.footerRight}>
          <div className={styles.socialIcons}>
            {linkedInUrl && (
              <a href={linkedInUrl} target="_blank" rel="noreferrer" className={styles.socialIcon} aria-label="LinkedIn">
                <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="20" height="20" fill="currentColor">
                  <path d="M20.447 20.452h-3.554v-5.569c0-1.328-.027-3.037-1.852-3.037-1.853 0-2.136 1.445-2.136 2.939v5.667H9.351V9h3.414v1.561h.046c.477-.9 1.637-1.85 3.37-1.85 3.601 0 4.267 2.37 4.267 5.455v6.286zM5.337 7.433a2.062 2.062 0 01-2.063-2.065 2.064 2.064 0 112.063 2.065zm1.782 13.019H3.555V9h3.564v11.452zM22.225 0H1.771C.792 0 0 .774 0 1.729v20.542C0 23.227.792 24 1.771 24h20.451C23.2 24 24 23.227 24 22.271V1.729C24 .774 23.2 0 22.222 0h.003z"/>
                </svg>
              </a>
            )}
            {youTubeUrl && (
              <a href={youTubeUrl} target="_blank" rel="noreferrer" className={styles.socialIcon} aria-label="YouTube">
                <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="20" height="20" fill="currentColor">
                  <path d="M23.495 6.205a3.007 3.007 0 00-2.088-2.088c-1.87-.501-9.396-.501-9.396-.501s-7.507-.01-9.396.501A3.007 3.007 0 00.527 6.205a31.247 31.247 0 00-.522 5.805 31.247 31.247 0 00.522 5.783 3.007 3.007 0 002.088 2.088c1.868.502 9.396.502 9.396.502s7.506 0 9.396-.502a3.007 3.007 0 002.088-2.088 31.247 31.247 0 00.5-5.783 31.247 31.247 0 00-.5-5.805zM9.609 15.601V8.408l6.264 3.602z"/>
                </svg>
              </a>
            )}
          </div>
          <p className={styles.copyright}>© {currentYear}</p>
        </div>

      </div>
    </footer>
  );
};

export default PageFooter;