import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'MyTodoWebPartStrings';
import MyTodo from './components/MyTodo';
import { IMyTodoProps } from './components/IMyTodoProps';
import { MSGraphClient } from '@microsoft/sp-http';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';

export interface IMyTodoWebPartProps {
  title: string;
  description: string;
  nrOfTasks: number;
  refreshInterval: number;
}

export default class MyTodoWebPart extends BaseClientSideWebPart<IMyTodoWebPartProps> {
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
    const element: React.ReactElement<IMyTodoProps> = React.createElement(
      MyTodo,
      {
        title: this.properties.title,
        refreshInterval: this.properties.refreshInterval,
        nrOfTasks: this.properties.nrOfTasks,
        description: this.properties.description,
        graphClient: this.graphClient,
        displayMode: this.displayMode,
        updateProperty: (value: string): void => {
          // store the new title in the title web part property
          this.properties.title = value;
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyFieldNumber("nrOfTasks", {
                  key: "nrOfTasks",
                  label: strings.NumberofTasks,
                  value: this.properties.nrOfTasks,
                  minValue: 1,
                  maxValue: 10
                }),
                PropertyFieldNumber("refreshInterval", {
                  key: "refreshInterval",
                  label: strings.RefreshInterval,
                  value: this.properties.refreshInterval,
                  minValue: 1,
                  maxValue: 60
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
