declare interface IPageHeaderWebPartStrings {
  PropertyPaneHeader: string;
  GroupNameGeneral: string;
  FieldSiteTitle: string;
  FieldSearchPlaceholder: string;
  FieldSearchPageUrl: string;
  FieldBackgroundImage: string;
  GroupNameSeasonal: string;
  FieldSeasonalEnabled: string;
  FieldSeasonalLabel: string;
  FieldSeasonalUrl: string;
}

declare module 'PageHeaderWebPartStrings' {
  const strings: IPageHeaderWebPartStrings;
  export = strings;
}