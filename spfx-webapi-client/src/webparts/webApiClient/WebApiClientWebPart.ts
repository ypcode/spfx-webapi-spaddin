import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'WebApiClientWebPartStrings';
import WebApiClient from './components/WebApiClient';
import { IWebApiClientProps } from './components/IWebApiClientProps';

export interface IWebApiClientWebPartProps {
  description: string;
}

export default class WebApiClientWebPart extends BaseClientSideWebPart<IWebApiClientWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IWebApiClientProps > = React.createElement(
      WebApiClient,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
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
