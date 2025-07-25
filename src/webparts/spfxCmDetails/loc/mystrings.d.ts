declare interface ISpfxCmDetailsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
    UnknownEnvironment: string;
    ApplicationDeadline: string;
    JobType: string;
    OpportunityDetails: string;
    ProgramArea: strings;
    Department: strings;
    NumberOpportunities: string;
    Duration: string;
    WorkArrangement: string;
    Location: string;
    SecurityLevel: string;
    LanguageRequirements: string;
    Apply: string;
    Edit: string;
    Retention: string;
    Expired: string;
    ApplicationsClosed: string;
    Delete: string;
    oppDeletedTitle: string;
    oppDeletedSubText: string;
    contactUs: string;
    cmHomePage: string;
    loading: string;
    dialogTitle: string;
    dialogText: string;
    cancel: string;
    emailSubject: string;
}

declare module 'SpfxCmDetailsWebPartStrings' {
  const strings: ISpfxCmDetailsWebPartStrings;
  export = strings;
}
