// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneChoiceGroupOption,
  PropertyPaneChoiceGroup,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import {BaseClientSideWebPart} from '@microsoft/sp-webpart-base';
import * as strings from 'MyDayEmailWebPartStrings';
import { MyDayEmail, IMyDayEmailProps } from './components';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldToggleWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout';
import { MSGraphClient } from '@microsoft/sp-http';

export interface IMyDayEmailWebPartProps {
  title: string;
  nrOfMessages: number;  
  refreshInterval: number;
  emailType: string;
}

export default class MyDayEmailWebPart extends BaseClientSideWebPart<IMyDayEmailWebPartProps> {
  private graphClient: MSGraphClient;

  public onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          this.graphClient = client;
          resolve();
        }, err => reject(err));
    });
  }

  public render(): void {
    const element: React.ReactElement<IMyDayEmailProps> = React.createElement(
      MyDayEmail,
      {
        title: this.properties.title,
        refreshInterval: this.properties.refreshInterval,
        nrOfMessages: this.properties.nrOfMessages,
        emailType: this.properties.emailType,
        // pass the current display mode to determine if the title should be
        // editable or not
        displayMode: this.displayMode,
        // pass the reference to the MSGraphClient
        graphClient: this.graphClient,
        // handle updated web part title
        updateProperty: (value: string): void => {
          // store the new title in the title web part property
          this.properties.title = value;
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    // const emailTypes: IPropertyPaneChoiceGroupOption[] = [];
    // ["Default", strings.FlaggedPivot, strings.ImportantPivot].map((item) => {
    //   emailTypes.push({key: item, text: item});
    // });

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupFields: [
                PropertyFieldNumber("nrOfMessages", {
                  key: "nrOfMessages",
                  label: strings.NrOfMessagesToShow,
                  value: this.properties.nrOfMessages,
                  minValue: 1,
                  maxValue: 10
                }),
                PropertyFieldNumber("refreshInterval", {
                  key: "refreshInterval",
                  label: strings.RefreshInterval,
                  value: this.properties.refreshInterval,
                  minValue: 1,
                  maxValue: 60
                })//,
                // PropertyPaneChoiceGroup("emailType", {
                //   label: strings.EmailTypePropertyLabel,
                //   options: emailTypes
                // })
              ]
            }
          ]
        }
      ]
    };
  }
}
