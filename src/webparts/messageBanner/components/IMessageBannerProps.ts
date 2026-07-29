import { SPFI } from '@pnp/sp';

export interface IMessageBannerProps {
  sp: SPFI;
  isFrench: boolean;
  onNoItems: () => void;
  onHasItems: () => void;
}