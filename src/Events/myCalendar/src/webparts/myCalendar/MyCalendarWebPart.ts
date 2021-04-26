// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';

import * as strings from 'MyCalendarWebPartStrings';
import MyCalendar from './components/MyCalendar';
import { IMyCalendarProps } from './components/IMyCalendarProps';
import { MSGraphClient } from '@microsoft/sp-http';
import { Providers, SharePointProvider } from '@microsoft/mgt';

export interface IMyCalendarWebPartProps {
  title: string;
  refreshInterval: number;
  daysInAdvance: number;
  numMeetings: number;
}

export default class MyCalendarWebPart extends BaseClientSideWebPart<IMyCalendarWebPartProps> {
  private propertyFieldNumber;

  public onInit(): Promise<void> {
    Providers.globalProvider = new SharePointProvider(this.context);
    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<IMyCalendarProps> = React.createElement(
      MyCalendar,
      {
        title: this.properties.title,
        refreshInterval: this.properties.refreshInterval,
        daysInAdvance: this.properties.daysInAdvance,
        numMeetings: this.properties.numMeetings,
        // pass the current display mode to determine if the title should be
        // editable or not
        displayMode: this.displayMode,
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

  //executes only before property pane is loaded.
  protected async loadPropertyPaneResources(): Promise<void> {
    // import additional controls/components

    const { PropertyFieldNumber } = await import(
      /* webpackChunkName: 'pnp-propcontrols-number' */
      '@pnp/spfx-property-controls/lib/propertyFields/number'
    );

    this.propertyFieldNumber = PropertyFieldNumber;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupFields: [
                this.propertyFieldNumber("refreshInterval", {
                  key: "refreshInterval",
                  // label: strings.RefreshInterval,
                  label: strings.RefreshInterval,
                  value: this.properties.refreshInterval,
                  minValue: 1,
                  maxValue: 60
                }),
                PropertyPaneSlider('daysInAdvance', {
                  label: strings.DaysInAdvance,
                  min: 0,
                  max: 7,
                  step: 1,
                  value: this.properties.daysInAdvance
                }),
                PropertyPaneSlider('numMeetings', {
                  label: strings.NumMeetings,
                  min: 0,
                  max: 20,
                  step: 1,
                  value: this.properties.numMeetings
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
