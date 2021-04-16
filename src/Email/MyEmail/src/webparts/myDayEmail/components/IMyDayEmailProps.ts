// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { IMyDayEmailWebPartProps } from "../MyDayEmailWebPart";
import { MSGraphClient } from "@microsoft/sp-http";
import { DisplayMode } from "@microsoft/sp-core-library";

export interface IMyDayEmailProps extends IMyDayEmailWebPartProps {
  displayMode: DisplayMode;
  graphClient: MSGraphClient;
  updateProperty: (value: string) => void;
}
