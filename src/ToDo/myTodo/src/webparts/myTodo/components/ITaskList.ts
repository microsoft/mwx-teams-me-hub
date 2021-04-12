export interface ITasksLists {
    value: ITaskList[];
  }
  
  export interface ITaskList {
    id:string;
    displayName:string;
    wellknownListName:string;
  }