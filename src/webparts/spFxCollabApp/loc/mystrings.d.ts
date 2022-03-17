declare interface ISpFxCollabAppWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'SpFxCollabAppWebPartStrings' {
  const strings: ISpFxCollabAppWebPartStrings;
  export = strings;
}
