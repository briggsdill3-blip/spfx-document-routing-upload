declare interface ISiteContentIndexWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  ConfigurationGroupName: string;
  TargetSitesFieldLabel: string;
  IncludeSystemListsFieldLabel: string;
  GroupBySiteFieldLabel: string;
  OnlyUniquePermissionsFieldLabel: string;
  StaleDaysThresholdFieldLabel: string;
  TitleGroupName: string;
  ShowTitleFieldLabel: string;
  CustomTitleFieldLabel: string;
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

declare module 'SiteContentIndexWebPartStrings' {
  const strings: ISiteContentIndexWebPartStrings;
  export = strings;
}