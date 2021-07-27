import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ClaringtonStaffDirectoryWebPartStrings';
import ClaringtonStaffDirectory from './components/ClaringtonStaffDirectory';
import { IClaringtonStaffDirectoryProps } from './components/IClaringtonStaffDirectory';
import { ClientMode } from './components/ClientMode';
import { IUser } from './interface/IUser';

export interface IClaringtonStaffDirectoryWebPartProps {
  description: string;
  clientMode: ClientMode;
  users: IUser[];
}

export default class ClaringtonStaffDirectoryWebPart extends BaseClientSideWebPart<IClaringtonStaffDirectoryWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IClaringtonStaffDirectoryProps> = React.createElement(
      ClaringtonStaffDirectory,
      {
        clientMode: this.properties.clientMode,
        description: this.properties.description,
        users: [],
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
