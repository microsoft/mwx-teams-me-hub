// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { MSGraphClient } from "@microsoft/sp-http";
import { DisplayMode } from "@microsoft/sp-core-library";
import { IMyTodoWebPartProps } from "../MyTodoWebPart";

export interface IMyTodoProps extends IMyTodoWebPartProps {
  displayMode: DisplayMode;
  description: string;
  graphClient: MSGraphClient;
  updateProperty: (value: string) => void;
}
