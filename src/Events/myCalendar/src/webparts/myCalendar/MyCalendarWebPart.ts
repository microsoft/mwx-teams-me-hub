// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { HeaderDisplay, ClickAction } from './enums';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';

import * as strings from 'MyCalendarWebPartStrings';
import MyCalendar from './components/MyCalendar';
import { IMyCalendarProps } from './components/IMyCalendarProps';
import { MSGraphClient } from '@microsoft/sp-http';


export interface IMyCalendarWebPartProps {
  title: string;
  refreshInterval: number;
  daysInAdvance: number;
  numMeetings: number;
  clickAction: ClickAction;
  showNew: boolean;
  showViewAll: boolean;
  headerDisplay?: HeaderDisplay;
}

export default class MyCalendarWebPart extends BaseClientSideWebPart<IMyCalendarWebPartProps> {
  private propertyFieldNumber;
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
    const element: React.ReactElement<IMyCalendarProps> = React.createElement(
      MyCalendar,
      {
        title: this.properties.title,
        refreshInterval: this.properties.refreshInterval,
        daysInAdvance: this.properties.daysInAdvance,
        numMeetings: this.properties.numMeetings,
        clickAction: this.properties.clickAction,
        showNew: this.properties.showNew,
        showViewAll: this.properties.showViewAll,
        headerDisplay: this.properties.headerDisplay,
        // pass the current display mode to determine if the title should be
        // editable or not
        displayMode: this.displayMode,
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
              groupName: strings.behaviorPropertyGroupName,

              groupFields: [
                // this.propertyFieldNumber("refreshInterval", {
                //   key: "refreshInterval",
                //   // label: strings.RefreshInterval,
                //   label: strings.RefreshIntervalPropertyLabel,
                //   value: this.properties.refreshInterval,
                //   minValue: 1,
                //   maxValue: 60
                // }),
                PropertyPaneSlider("refreshInterval", {                  
                  label: strings.RefreshIntervalPropertyLabel,
                  value: this.properties.refreshInterval,
                  min: 1,
                  max: 60,
                  step: 1

                }),
                PropertyPaneSlider('daysInAdvance', {
                  label: strings.DaysInAdvancePropertyLabel,
                  min: 0,
                  max: 7,
                  step: 1,
                  value: this.properties.daysInAdvance
                }),
                PropertyPaneSlider('numMeetings', {
                  label: strings.NumMeetingsPropertyLabel,
                  min: 0,
                  max: 20,
                  step: 1,
                  value: this.properties.numMeetings
                }),
                PropertyPaneDropdown("clickAction", {
                  label: strings.ClickActionPropertyLabel,
                  ariaLabel: strings.ClickActionPropertyLabel,
                  selectedKey: this.properties.clickAction,                  
                  options: [
                    {                      
                      text: ClickAction.Preview,
                      key: ClickAction.Preview
                    },                    
                    {                     
                      text: ClickAction.OpenInOutlook,
                      key: ClickAction.OpenInOutlook
                    }
                  ]
                })
              ]
            },            
            {
              groupName: strings.capabilitiesPropertyGroupName,
              
              groupFields: [                
                
                // PropertyPaneToggle("enableDelete", {
                //   label: strings.EnableDeletePropertyLabel,
                //   key: "enableDelete",
                //   onText: strings.EnableOnPropertyText,
                //   onAriaLabel: strings.EnableOnPropertyText,
                //   offText: strings.EnableOffPropertyText,
                //   offAriaLabel: strings.EnableOffPropertyText,
                //   checked: this.properties.enableDelete
                // }),
                // PropertyPaneToggle("enableToggleReadStatus", {
                //   label: strings.EnableToggleReadPropertyLabel,
                //   key: "enableToggleReadStatus",
                //   onText: strings.EnableOnPropertyText,
                //   onAriaLabel: strings.EnableOnPropertyText,
                //   offText: strings.EnableOffPropertyText,
                //   offAriaLabel: strings.EnableOffPropertyText,
                //   checked: this.properties.enableToggleReadStatus
                // }),
                // PropertyPaneToggle("enableToggleFlag", {
                //   label: strings.EnableToggleFlagPropertyLabel,
                //   key: "enableToggleFlag",
                //   onText: strings.EnableOnPropertyText,
                //   onAriaLabel: strings.EnableOnPropertyText,
                //   offText: strings.EnableOffPropertyText,
                //   offAriaLabel: strings.EnableOffPropertyText,
                //   checked: this.properties.enableToggleFlag
                // }),
                PropertyPaneToggle("showNew", {
                  label: strings.ShowNewPropertyLabel,
                  key: "showNew",
                  onText: strings.ShowOnPropertyText,
                  onAriaLabel: strings.ShowOnPropertyText,
                  offText: strings.ShowOffPropertyText,
                  offAriaLabel: strings.ShowOffPropertyText,
                  checked: this.properties.showNew
                }),
                PropertyPaneToggle("showViewAll", {
                  label: strings.ShowViewAllPropertyLabel,
                  key: "showViewAll",
                  onText: strings.ShowOnPropertyText,
                  onAriaLabel: strings.ShowOnPropertyText,
                  offText: strings.ShowOffPropertyText,
                  offAriaLabel: strings.ShowOffPropertyText,
                  checked: this.properties.showViewAll
                })
              ],
              
            },
            {
              groupName: strings.uiPropertyGroupName,
              
              groupFields: [                
                PropertyPaneDropdown("headerDisplay", {
                  label: strings.HeaderDisplayPropertyLabel,
                  ariaLabel: strings.HeaderDisplayPropertyLabel,
                  selectedKey: this.properties.headerDisplay,                  
                  options: [
                    {                      
                      text:  HeaderDisplay.Standard,
                      key: HeaderDisplay.Standard
                    },
                    {                      
                      text: HeaderDisplay.Large,
                      key: HeaderDisplay.Large
                    },
                    {                     
                      text: HeaderDisplay.None,
                      key: HeaderDisplay.None
                    }
                  ]
                })/*,
              
                PropertyPaneToggle("enableThemes", {
                  label: strings.EnableThemesPropertyLabel,
                  key: "enableThemes",
                  onText: strings.EnableOnPropertyText,
                  onAriaLabel: strings.EnableOnPropertyText,
                  offText: strings.EnableOffPropertyText,
                  offAriaLabel: strings.EnableOffPropertyText,
                  checked: this.properties.enableThemes
                })*/
              ]
            }
          ]
        }
      ]
    };
  }
}
