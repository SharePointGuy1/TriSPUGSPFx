import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'helloTriSpugStrings';
import HelloTriSpug from './components/HelloTriSpug';
import { IHelloTriSpugProps } from './components/IHelloTriSpugProps';
import { IHelloTriSpugWebPartProps } from './IHelloTriSpugWebPartProps';

export default class HelloTriSpugWebPart extends BaseClientSideWebPart<IHelloTriSpugWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IHelloTriSpugProps > = React.createElement(
      HelloTriSpug,
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
