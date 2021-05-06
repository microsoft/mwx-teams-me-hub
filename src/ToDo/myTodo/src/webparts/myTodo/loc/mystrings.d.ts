// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
declare interface IMyTodoWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  Error: string;
  Loading: string;
  NumberofTasks: string;
  NoMessages: string;
  Loading: string;
  RefreshInterval:string;
  ViewAllTodo:string;
}

declare module 'MyTodoWebPartStrings' {
  const strings: IMyTodoWebPartStrings;
  export = strings;
  
}
