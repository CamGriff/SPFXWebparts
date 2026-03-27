declare interface IPageHeaderWebPartStrings {
  PropertyPaneHeader: string;
  GroupNameGeneral: string;
  FieldSearchPlaceholder: string;
  FieldSearchPageUrl: string;
  FieldBackgroundImage: string;
  FieldSiteTitle: string;
}

declare module 'PageHeaderWebPartStrings' {
  const strings: IPageHeaderWebPartStrings;
  export = strings;
}