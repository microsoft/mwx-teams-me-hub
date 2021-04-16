// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { ITaskList } from './ITaskList';
import { ITask } from './ITasks';

export interface IMyTodoState {
  tasks: ITask[];
  renderedDateTime: Date;
  error: string;
  loading: boolean;
  activeTaskList: ITaskList;
}