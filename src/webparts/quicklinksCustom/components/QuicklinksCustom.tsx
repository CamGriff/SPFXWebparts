import * as React from 'react';
import styles from './QuicklinksCustom.module.scss';
import type { IQuicklinksCustomProps, IQuickLinkTile } from './IQuicklinksCustomProps';
import { Icon } from '@fluentui/react/lib/Icon';

const QuicklinksCustom: React.FC<IQuicklinksCustomProps> = ({ tiles }) => {

  if (!tiles || tiles.length === 0) {
    return (
      <div className={styles.emptyState}>
        <p>No tiles configured. Edit this web part to add quick links.</p>
      </div>
    );
  }

  return (
    <div className={styles.quicklinksCustom}>
      {tiles.map((tile: IQuickLinkTile, index: number) => (
  <a 
    key={index}
    href={tile.url}
    className={styles.tile}
    target="_self"
    aria-label={tile.title}
  >
    {tile.icon && (
      <div className={styles.tileIcon}>
        <Icon iconName={tile.icon} />
      </div>
    )}
    <div className={styles.tileContent}>
      <div className={styles.tileTitle}>{tile.title}</div>
      {tile.description && (
        <div className={styles.tileDescription}>{tile.description}</div>
      )}
    </div>
  </a>
))}
    </div>
  );
};

export default QuicklinksCustom;