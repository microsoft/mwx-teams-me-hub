// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
declare interface IMyDayEmailWebPartStrings {
  // Property Pane
  PropertyPaneDescription: string;  
  HeaderDisplayPropertyLabel: string;
  MessagesToShowPropertyLabel: string;
  RefreshIntervalPropertyLabel:string;
  EmailDisplayPropertyLabel: string;
  ClickActionPropertyLabel: string;
  ShowNewPropertyLabel: string;  
  ShowViewAllPropertyLabel: string;
  EnableDeletePropertyLabel: string;
  EnableToggleFlagPropertyLabel: string;
  EnableToggleReadPropertyLabel: string;
  EnableThemesPropertyLabel: string;
  ShowOnPropertyText: string;
  ShowOffPropertyText: string;
  EnableOnPropertyText: string;
  EnableOffPropertyText: string;
  behaviorPropertyGroupName: string;
  capabilitiesPropertyGroupName: string;
  uiPropertyGroupName: string;

  // Application
  Error: string;
  Loading: string;
  NewEmail: string;
  NoMessages: string;    
  ViewAll: string;
  AllPivot: string;
  UnreadPivot: string;
  // FlaggedPivot: string;
  // ImportantPivot: string;  
}

declare module 'MyDayEmailWebPartStrings' {
  const strings: IMyDayEmailWebPartStrings;
  export = strings;
}
