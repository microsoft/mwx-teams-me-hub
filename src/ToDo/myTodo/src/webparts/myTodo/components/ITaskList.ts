// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
export interface ITasksLists {
    value: ITaskList[];
  }
  
  export interface ITaskList {
    id:string;
    displayName:string;
    wellknownListName:string;
  }