import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IFooterLink {
  title: string;
  url: string;
  category: string;
  sortOrder: number;
}

export interface IPageFooterProps {
  footerListName: string;
  linkedInUrl: string;
  youTubeUrl: string;
  logoUrl: string;
  context: WebPartContext;
}