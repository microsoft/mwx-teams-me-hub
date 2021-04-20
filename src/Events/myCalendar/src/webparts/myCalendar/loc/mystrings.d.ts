// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
declare interface IMyCalendarWebPartStrings {
  AllDay: string;
  DaysInAdvance: string;
  Error: string;
  Hour: string;
  Hours: string;
  Loading: string;
  Minutes: string;
  NewMeeting: string;
  NoMeetings: string;
  NumMeetings: string;
  RefreshInterval: string;
  PropertyPaneDescription: string;
  ViewAll: string;
}

declare module 'MyCalendarWebPartStrings' {
  const strings: IMyCalendarWebPartStrings;
  export = strings;
}
