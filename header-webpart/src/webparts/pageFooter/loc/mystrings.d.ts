declare interface IPageFooterWebPartStrings {
  PropertyPaneHeader: string;
  GroupNameGeneral: string;
  GroupNameSocial: string;
  FieldFooterListName: string;
  FieldLinkedInUrl: string;
  FieldYouTubeUrl: string;
  FieldLogoImage: string;
}

declare module 'PageFooterWebPartStrings' {
  const strings: IPageFooterWebPartStrings;
  export = strings;
}
