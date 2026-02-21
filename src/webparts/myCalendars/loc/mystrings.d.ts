declare interface IMyCalendarsWebPartStrings {
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
  CalendarSettingsTitle: string;
  CalendarSourcesLabel: string;
  AddCalendarLabel: string;
  AddCalendarSharePointLabel: string;
  AddCalendarSharePointDescription: string;
  AddCalendarExchangeLabel: string;
  AddCalendarExchangeDescription: string;
  AddCalendarIcsLabel: string;
  AddCalendarIcsDescription: string;
  SearchSitesLabel: string;
  SelectSiteLabel: string;
  SelectListLabel: string;
  CalendarNameLabel: string;
  CalendarColorLabel: string;
  FieldTitleCandidates: string;
  FieldStartCandidates: string;
  FieldEndCandidates: string;
  SelectMailboxLabel: string;
  EnterMailboxLabel: string;
  MailboxPlaceholder: string;
  SelectCalendarLabel: string;
  DoneLabel: string;
  DeleteLabel: string;
  BackLabel: string;
  SearchLabel: string;
  LoadCalendarsLabel: string;
  NoCalendarsFoundLabel: string;
  LoadingLabel: string;
}

declare module 'MyCalendarsWebPartStrings' {
  const strings: IMyCalendarsWebPartStrings;
  export = strings;
}
