declare interface IHelloTriSpugStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  PurposeFieldLabel: string;
  ShowHiddenFieldLabel: string;
  HowManyFieldLabel: string;
  ListTypeFieldLabel: string;
  ShowUrlFieldLabel: string;
  WebAddressFieldLabel: string;
}

declare module 'helloTriSpugStrings' {
  const strings: IHelloTriSpugStrings;
  export = strings;
}
