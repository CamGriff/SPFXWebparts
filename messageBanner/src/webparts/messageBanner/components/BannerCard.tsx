import * as React from 'react';
import { IBannerItem } from '../models/IBannerItem';
import styles from './MessageBanner.module.scss';

export interface IBannerCardProps {
  item: IBannerItem;
  isFrench: boolean;
}

const BannerCard: React.FC<IBannerCardProps> = ({ item, isFrench }) => {
  const title = isFrench ? item.TitleFR || item.Title : item.Title;
  const description = isFrench ? item.DescriptionFR || item.Description : item.Description;
  const seeMoreUrl = isFrench ? item.SeeMoreFRUrl || item.SeeMoreUrl : item.SeeMoreUrl;
  const seeMoreLabel = isFrench
    ? item.SeeMoreFRDescription || item.SeeMoreDescription || 'Voir plus'
    : item.SeeMoreDescription || 'See more';
  const levelStyleMap: Record<string, string> = {
  low: styles.low,
  medium: styles.medium,
  high: styles.high,
  };

const levelStyle = levelStyleMap[item.MessageLevel.toLowerCase()] ?? '';

return (
  <div className={`${styles.card} ${levelStyle}`}>
      <div className={styles.leftBorder} />
       <div className={styles.content}>
        <div className={styles.topRow}>
          <span className={styles.levelBadge}>{item.MessageLevel}</span>
          <span className={styles.title}>{title}</span>
        </div>
        {description && (
          <p className={styles.description}>{description}</p>
        )}
      </div>
      {seeMoreUrl && (
        <a
          href={seeMoreUrl}
          target="_blank"
          rel="noopener noreferrer"
          className={styles.seeMore}
        >
          {seeMoreLabel} {'>'}
        </a>
      )}
    </div>
  );
};

export default BannerCard;