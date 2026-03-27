import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IPageHeaderProps {
  siteTitle: string;
  searchBoxPlaceholder: string;
  searchPageUrl: string;
  backgroundImageUrl: string;
  isDarkTheme: boolean;
  context: WebPartContext;
}