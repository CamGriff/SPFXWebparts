import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { IMessageBannerProps } from './IMessageBannerProps';
import { BannerService } from '../services/BannerService';
import { IBannerItem } from '../models/IBannerItem';
import BannerCard from './BannerCard';
import styles from './MessageBanner.module.scss';

const FLIP_INTERVAL_MS = 8000;

const MessageBanner: React.FC<IMessageBannerProps> = ({ sp, isFrench, onNoItems, onHasItems }) => {
  const [items, setItems] = useState<IBannerItem[]>([]);
  const [currentIndex, setCurrentIndex] = useState(0);
  const [flipping, setFlipping] = useState(false);
  const [loading, setLoading] = useState(true);
  const intervalRef = useRef<ReturnType<typeof setInterval> | null>(null);
  

useEffect(() => {
  const service = new BannerService(sp);
  service.getActiveItems()
    .then(data => {
      setItems(data);
      setLoading(false);
      if (data.length > 0) onHasItems();
    })
    .catch(() => setLoading(false));
}, []);

const triggerFlip = (): void => {
    setFlipping(true);
    setTimeout(() => {
      setCurrentIndex(prev => (prev + 1) % items.length);
      setFlipping(false);
    }, 400); // half of CSS animation duration
  };

  useEffect(() => {
    if (items.length <= 1) return;

    intervalRef.current = setInterval(() => {
      triggerFlip();
    }, FLIP_INTERVAL_MS);

    return () => {
      if (intervalRef.current) clearInterval(intervalRef.current);
    };
  }, [items, currentIndex]);

  if (loading) return null;
  if (items.length === 0) {
  onNoItems();
  return null;
  }

  return (
    <div className={styles.wrapper}>
      <div className={`${styles.flipContainer} ${flipping ? styles.flipping : ''}`}>
        <BannerCard item={items[currentIndex]} isFrench={isFrench} />
      </div>
      {items.length > 1 && (
        <div className={styles.dots}>
          {items.map((_, i) => (
            <span
              key={i}
              className={`${styles.dot} ${i === currentIndex ? styles.activeDot : ''}`}
            />
          ))}
        </div>
      )}
    </div>
  );
};

export default MessageBanner;