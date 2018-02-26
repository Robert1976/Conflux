declare interface ICswpWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  QueryFieldLabel: string;
  ItemTemplateFieldLabel: string;
  ControlTemplateFieldLabel: string;
  CssStylesFieldLabel: string;
  MaximumNumberResultsFieldLabel: string;
  NoResultsTemplateFieldLabel: string;
}

declare module 'CswpWebPartStrings' {
  const strings: ICswpWebPartStrings;
  export = strings;
}
