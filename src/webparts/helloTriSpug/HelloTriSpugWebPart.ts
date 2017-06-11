import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneLink
} from '@microsoft/sp-webpart-base';

import * as strings from 'helloTriSpugStrings';
import HelloTriSpug from './components/HelloTriSpug';
import { IHelloTriSpugProps } from './components/IHelloTriSpugProps';
import { IHelloTriSpugWebPartProps } from './IHelloTriSpugWebPartProps';

export default class HelloTriSpugWebPart extends BaseClientSideWebPart<IHelloTriSpugWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IHelloTriSpugProps> = React.createElement(
      HelloTriSpug,
      {
        description: this.properties.description,
        purpose: this.properties.purpose,
        showHidden: this.properties.showHidden,
        howMany: this.properties.howMany,
        listType: this.properties.listType,
        showUrl: this.properties.showUrl,
        url: this.properties.url
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
                }),
                PropertyPaneTextField('purpose', {
                  label: 'What is the purpose of this webpart?',
                  multiline: true
                }),
                PropertyPaneCheckbox('showHidden', {
                  text: strings.ShowHiddenFieldLabel
                }),
                PropertyPaneTextField('howMany', { label: 'Show how many lists' }),
                PropertyPaneDropdown('listType', {
                  label: strings.ListTypeFieldLabel,
                  options: [
                    { key: 'All', text: 'All' },
                    { key: '0', text: 'Lists Only' },
                    { key: '1', text: 'Document Libraries' }
                  ]
                }),
                PropertyPaneToggle('showUrl', {
                  label: 'Show the url',
                  onText: 'Show',
                  offText: "Don't show"
                }),
                PropertyPaneLink('url', {
                  target: "new",
                  href: 'https://holmesinfosys.sharepoint.com/sites/MyDevSite',
                  text: 'My Site'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
