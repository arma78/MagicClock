import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {SPHttpClient} from "@microsoft/sp-http";
import * as strings from 'MsTeamsClockWebPartStrings';
import MsTeamsClock from './components/MsTeamsClock';
import { IMsTeamsClockProps } from './components/IMsTeamsClockProps';

export interface IMsTeamsClockWebPartProps {
  description: string;
  title: string;
  spHttpClient: SPHttpClient;
}

export default class MsTeamsClockWebPart extends BaseClientSideWebPart<IMsTeamsClockWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IMsTeamsClockProps> = React.createElement(
      MsTeamsClock,
      {
        description: this.properties.description,
        title: this.properties.title,
        context: this.context
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
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
