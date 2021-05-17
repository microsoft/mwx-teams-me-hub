// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { IMyDayEmailWebPartProps } from "../MyDayEmailWebPart";
import { MSGraphClient } from "@microsoft/sp-http";
import { DisplayMode } from "@microsoft/sp-core-library";

export interface IMyDayEmailProps extends IMyDayEmailWebPartProps {
  displayMode: DisplayMode;
  graphClient: MSGraphClient;  
  themeVariant: IReadonlyTheme | undefined;
  updateProperty: (value: string) => void;
}
