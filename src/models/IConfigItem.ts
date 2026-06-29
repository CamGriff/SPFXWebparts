export interface IConfigItem {
  Id?: number;
  Title: string;
  SiteUrl: string;
  Language: string;
  StaleDays: number;
  TranslatorName: string;
  TranslatorEmail: string;
  IsActive: boolean;
}
