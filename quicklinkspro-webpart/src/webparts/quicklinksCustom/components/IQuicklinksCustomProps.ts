import { DisplayMode } from '@microsoft/sp-core-library';

export interface IQuickLinkTile {
  title: string;
  description: string;
  url: string;
  icon: string;
}

export interface IQuicklinksCustomProps {
  tiles: IQuickLinkTile[];
  isDarkTheme: boolean;
  displayMode: DisplayMode;
}