// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { IMyCalendarWebPartProps } from "../MyCalendarWebPart";
import { DisplayMode } from "@microsoft/sp-core-library";
import { MSGraphClient } from "@microsoft/sp-http";

export interface IMyCalendarProps extends IMyCalendarWebPartProps {
  displayMode: DisplayMode;
  graphClient: MSGraphClient;
  updateProperty: (value: string) => void;
}
