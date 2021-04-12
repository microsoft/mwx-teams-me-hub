export interface ITasks {
    value: ITask[];
  }
  
  export interface ITask {
    title: string;
    isReminderOn:boolean;
    status:string;
    createdDateTime:string;
    importance:string;
    id:string;
    dueDateTime:{
      dateTime:string;
      timeZone:string;
    };
  }