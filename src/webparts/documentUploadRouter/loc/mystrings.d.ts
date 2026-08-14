declare interface IDocumentUploadRouterWebPartStrings {
  PropertyPaneDescription: string;
  TitleGroupName: string;
  ConfigurationGroupName: string;
  ShowTitleFieldLabel: string;
  CustomTitleFieldLabel: string;
  TargetSitesFieldLabel: string;
  IncludeSystemListsFieldLabel: string;
  GroupBySiteFieldLabel: string;
  ExpandByDefaultFieldLabel: string;
  StaleDaysThresholdFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'DocumentUploadRouterWebPartStrings' {
  const strings: IDocumentUploadRouterWebPartStrings;
  export = strings;
}