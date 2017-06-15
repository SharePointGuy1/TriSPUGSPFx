import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library'; // there is separate import for Environment pieces
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneLink,
  PropertyPaneHorizontalRule,
  PropertyPaneCustomField,            // for sample custom field
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';

import * as strings from 'helloTriSpugStrings';
import HelloTriSpug from './components/HelloTriSpug';
import { IHelloTriSpugProps } from './components/IHelloTriSpugProps';
import { IHelloTriSpugWebPartProps } from './IHelloTriSpugWebPartProps';

import MockHttpClient from './MockHttpClient';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
  BaseType: number;
}

import styles from './components/HelloTriSpug.module.scss';

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
    if (this.displayMode == 1) { // do not update if in edit mode (https://github.com/SharePoint/sp-dev-docs/blob/master/reference/spfx/sp-core-library/displaymode.md)
      this._renderListAsync();
    }
  }

  // turns off all reactive property changes
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected _renderListAsync() {
    // Local environment
    if (Environment.type === EnvironmentType.Local) {
      this._getMockListData().then((response) => {
        this._renderList(response.value);
      });
    }
    else if (Environment.type == EnvironmentType.SharePoint ||
      Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getListData()
        .then((response) => {
          this._renderList(response.value);
        });
    }
  }
  private _getMockListData() {
    return MockHttpClient.get()
      .then((data: ISPList[]) => {
        var listData: ISPLists = { value: data };
        return listData;
      }) as Promise<ISPLists>;
  }

  private _getListData(): Promise<ISPLists> {
    let api: string = this.context.pageContext.web.absoluteUrl + `/_api/web/lists`;
    let filter: string = '';
    if (!this.properties.showHidden) {
      filter = '?$filter=Hidden eq false';
    }
    if (this.properties.howMany) {
      if (filter) {
        filter += "&";
      } else {
        filter = '?';
      }
      filter += '$top=' + this.properties.howMany.toString();
    }
    if (filter) {
      api += filter;
    }
    return this.context.spHttpClient.get(api, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _renderList(items: ISPList[]): void {

    let html: string = '';
    html += '<p class="ms-font-l ms-fontColor-blue">Loading from \'' + this.context.pageContext.web.title + '\'</p>\n';

    html += '<ul class="${styles.listItem}">';
    items.forEach((item: ISPList) => {
      html += `        
            <li class="${styles.listItem}" baseType="${item.BaseType}">
                <span class="ms-font-l">${item.Title}</span>
            </li>
        `;
    });
    html += '</ul>';
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // render the custom field
  private _customFieldRender(elem: HTMLElement): void {
    elem.innerHTML = '<div><h1>This is a custom field.</h1></div>';
  }

  // for the async list dropdown
  private _options: IPropertyPaneDropdownOption[];
  protected onInit(): Promise<void> {
    return this._getLists().then(lists => {
      this._options = lists.map(list => {
        return {
          key: list.Id,
          text: list.Title
        };
      });
    });
  }

  private _getLists(): Promise<any> {
    if (Environment.type === EnvironmentType.Local) {
      return new Promise<any>(resolve => {
        setTimeout(() => resolve([
          { Title: 'Mock List', Id: '1', BaseType: 0 },
          { Title: 'Mock List 2', Id: '2', BaseType: 1 },
          { Title: 'Mock List 3', Id: '3', BaseType: 0 }]),
          500);
      });
    }
    else {
      const url: string = this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`;
      return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          return response.json();
        }).then((json) => {
          return json.value;
        });
    }
  }  // end for the async list dropdown


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
                PropertyPaneHorizontalRule(),
                PropertyPaneTextField('howMany', { label: 'Show how many lists' }),
                PropertyPaneDropdown('listType', {    // https://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.client.basetype.aspx
                  label: strings.ListTypeFieldLabel,
                  options: [
                    { key: '-1', text: 'All' },
                    { key: '0', text: 'Lists Only' }, // GenericList
                    { key: '1', text: 'Document Libraries' }, // DocumentLibrary
                    { key: '2', text: 'Unused' },       // Unused
                    { key: '3', text: 'DiscussionBoard' },  // DiscussionBoard
                    { key: '4', text: 'Survey' }, // Survey
                    { key: '5', text: 'Issue' }
                  ]
                }),
                PropertyPaneToggle('showUrl', {
                  label: strings.ShowUrlFieldLabel,// 'Show the url', 
                  onText: 'Show',
                  offText: "Don't show"
                }),
                PropertyPaneLink('url', {
                  target: "_blank",
                  href: 'https://holmesinfosys.sharepoint.com/sites/MyDevSite',
                  text: strings.WebAddressFieldLabel,// 'My Site'                
                })

              ]
            }
          ]
        }, {
          header: {
            description: "This is page 2"
          },
          groups: [{
            groupName: "Group 1 of Page 2",
            groupFields: [
              PropertyPaneCustomField({
                key: 'customField',
                onRender: this._customFieldRender.bind(this)
              })
            ]
          }, {
            groupName: "Group 2 of Page 2", groupFields: [PropertyPaneDropdown('listName', {  // for the async list dropdown
              label: 'Select a list',
              selectedKey: this._options.length > 0 ? this._options[0].key : null,
              options: this._options
            })]
          }]
        }
      ]
    };
  }
}
