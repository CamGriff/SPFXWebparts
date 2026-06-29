import { DriftStatus } from './DriftStatus';

export interface IDriftItem {
  Title: string;
  DefaultPageTitle: string;
  DefaultPageUrl: string;
  DefaultPageModified: string;
  TranslationLanguage: string;
  TranslationPageUrl: string;
  TranslationModified: string | undefined;
  DaysDrift: number;
  DriftStatus: DriftStatus;
  TranslatorName: string | undefined;
  TranslatorEmail: string | undefined;
  SiteUrl: string;
  LastChecked: string;
  PageGuid: string;
  NudgeSent: boolean;
  NudgeDate: string | undefined;
}
