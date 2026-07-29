export interface IBannerItem {
  Id: number;
  Title: string;
  TitleFR: string;
  Description: string;
  DescriptionFR: string;
  SeeMoreUrl: string;
  SeeMoreDescription: string;
  SeeMoreFRUrl: string;
  SeeMoreFRDescription: string;
  MessageLevel: 'Low' | 'Medium' | 'High';
  MessageStatus: 'Active' | 'Inactive';
  ExpirationDate: string;
}