// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as strings from 'MyDayEmailWebPartStrings';
import { HeaderDisplay, EMailDisplay, ClickAction } from './enums';
import { MyDayEmail, IMyDayEmailProps } from './components';

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneChoiceGroupOption,
  PropertyPaneChoiceGroup,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import {BaseClientSideWebPart} from '@microsoft/sp-webpart-base';
import {
  ThemeProvider,
  ThemeChangedEventArgs,
  IReadonlyTheme,
  ISemanticColors
} from '@microsoft/sp-component-base';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldToggleWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout';
import { MSGraphClient } from '@microsoft/sp-http';
import { EmailDisplayPropertyLabel } from 'MyDayEmailWebPartStrings';

export interface IMyDayEmailWebPartProps {
  title: string;
  headerDisplay?: HeaderDisplay;
  numberOfMessages: number;  
  refreshInterval: number;
  emailDisplay: EMailDisplay;
  clickAction: ClickAction;
  showNew: boolean;
  showViewAll: boolean;
  enableDelete: boolean;
  enableToggleReadStatus: boolean;
  enableThemes: boolean;
  enableToggleFlag: boolean;
}

export default class MyDayEmailWebPart extends BaseClientSideWebPart<IMyDayEmailWebPartProps> {
  private _graphClient: MSGraphClient;  
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  public async onInit(): Promise<any> {
    await super.onInit();

    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();

    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);

    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          this._graphClient = client;
          resolve();
        }, err => reject(err));
    });
  }

  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }

  public render(): void {  
    const element: React.ReactElement<IMyDayEmailProps> = React.createElement(
      MyDayEmail,
      {
        title: this.properties.title,        
        refreshInterval: this.properties.refreshInterval,
        numberOfMessages: this.properties.numberOfMessages,
        emailDisplay: this.properties.emailDisplay,
        headerDisplay: this.properties.headerDisplay,
        clickAction: this.properties.clickAction,
        showNew: this.properties.showNew,
        showViewAll: this.properties.showViewAll,
        enableDelete: this.properties.enableDelete,
        enableToggleReadStatus: this.properties.enableToggleReadStatus,
        enableThemes: this.properties.enableThemes,
        enableToggleFlag: this.properties.enableToggleFlag,

        // pass the current them variant so the component can reder properly
        themeVariant: this._themeVariant,
        // pass the current display mode to determine if the title should be editable or not
        displayMode: this.displayMode,
        // pass the reference to the MSGraphClient
        graphClient: this._graphClient,
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
              groupName: strings.behaviorPropertyGroupName,
              
              groupFields: [
                PropertyFieldNumber("numberOfMessages", {
                  key: "numberOfMessages",
                  label: strings.MessagesToShowPropertyLabel,
                  value: this.properties.numberOfMessages,
                  minValue: 1,
                  maxValue: 10
                }),
                PropertyFieldNumber("refreshInterval", {
                  key: "refreshInterval",
                  label: strings.RefreshIntervalPropertyLabel,
                  value: this.properties.refreshInterval,
                  minValue: 1,
                  maxValue: 60
                }),
                PropertyPaneDropdown("emailDisplay", {
                  label: strings.EmailDisplayPropertyLabel,
                  ariaLabel: strings.EmailDisplayPropertyLabel,
                  selectedKey: this.properties.emailDisplay,                                    
                  options: [
                    {                      
                      text: EMailDisplay.Default,
                      key: EMailDisplay.Default
                    },
                    {                      
                      text: EMailDisplay.Flagged,
                      key: EMailDisplay.Flagged
                    },
                    {                     
                      text: EMailDisplay.Important,
                      key: EMailDisplay.Important
                    }
                  ]
                }),
                PropertyPaneDropdown("clickAction", {
                  label: strings.ClickActionPropertyLabel,
                  ariaLabel: strings.ClickActionPropertyLabel,
                  selectedKey: this.properties.clickAction,                  
                  options: [
                    {                      
                      text: ClickAction.PreviewUnread,
                      key: ClickAction.PreviewUnread
                    },
                    {                      
                      text: ClickAction.PreviewRead,
                      key: ClickAction.PreviewRead
                    },
                    {                     
                      text: ClickAction.OpenInOutlook,
                      key: ClickAction.OpenInOutlook
                    }
                  ]
                })
              ],
              
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
              ],
              
            }
          ]
        }
      ]
    };
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    console.log(propertyPath);
    
    if (propertyPath === 'emailDisplay') {
      
    }

    /*
    Check the property path to see which property pane feld changed. If the property path matches the dropdown, then we set that list
    as the selected list for the web part. 
    */
    // if (propertyPath === 'spListIndex') {
    //   this._setSelectedList(newValue);
    // }

    /*
    Finally, tell property pane to re-render the web part. 
    This is valid for reactive property pane. 
    */
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }
}
