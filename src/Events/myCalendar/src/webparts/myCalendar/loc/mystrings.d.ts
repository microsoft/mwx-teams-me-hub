// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
declare interface IMyCalendarWebPartStrings {
  PropertyPaneDescription: string;  
  HeaderDisplayPropertyLabel: string;
  MessagesToShowPropertyLabel: string;
  RefreshIntervalPropertyLabel:string; 
  DaysInAdvancePropertyLabel: string; 
  NumMeetingsPropertyLabel: string;
  ClickActionPropertyLabel: string;
  ShowNewPropertyLabel: string;  
  ShowViewAllPropertyLabel: string;  
  EnableThemesPropertyLabel: string;
  ShowOnPropertyText: string;
  ShowOffPropertyText: string;
  EnableOnPropertyText: string;
  EnableOffPropertyText: string;
  behaviorPropertyGroupName: string;
  capabilitiesPropertyGroupName: string;
  uiPropertyGroupName: string;
  
  AllDay: string;  
  Error: string;
  Hour: string;
  Hours: string;
  Loading: string;
  Minutes: string;
  NewMeeting: string;
  NoMeetings: string;    
  ViewAll: string;
}

declare module 'MyCalendarWebPartStrings' {
  const strings: IMyCalendarWebPartStrings;
  export = strings;
}
