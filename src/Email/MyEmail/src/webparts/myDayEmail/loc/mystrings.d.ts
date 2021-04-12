declare interface IMyDayEmailWebPartStrings {
  Error: string;
  Loading: string;
  NewEmail: string;
  NoMessages: string;
  NrOfMessagesToShow: string;
  PropertyPaneDescription: string;  
  ViewAll: string;
  RefreshInterval:string;
  AllPivot: string;
  UnreadPivot: string;
  FlaggedPivot: string;
  ImportantPivot: string;
  EmailTypePropertyLabel: string;
}

declare module 'MyDayEmailWebPartStrings' {
  const strings: IMyDayEmailWebPartStrings;
  export = strings;
}
