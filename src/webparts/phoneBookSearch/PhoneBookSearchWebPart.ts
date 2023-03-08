import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-webpart-base';

import PhoneBookSearch from './components/PhoneBookSearch';
import { IPhoneBookSearchProps } from './components/IPhoneBookSearchProps';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IPhoneBookSearchWebPartProps {
  wpTitle: string;
  listData: any;
  position: string;
  context: any;
}

export default class PhoneBookSearchWebPart extends BaseClientSideWebPart<IPhoneBookSearchWebPartProps> {

  private _getListData(): Promise<any> {
    return this.context.spHttpClient.get(
      this.context.pageContext.web.absoluteUrl +
      `/_api/web/lists/GetByTitle('PhoneBook')/Items?$filter=Aktiv%20eq%201&$top=5000`,
      SPHttpClient.configurations.v1
    ).then((result: SPHttpClientResponse) => {
      return result.json();
    });
  }

  public onInit(): Promise<void> {
    return this._getListData().then((response) => {
      return this.properties.listData = response.value;
    });
  }

  public render(): void {
    const element: React.ReactElement<IPhoneBookSearchProps> = React.createElement(
      PhoneBookSearch,
      {
        wpTitle: this.properties.wpTitle,
        listData: this.properties.listData,
        position: this.properties.position,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // @ts-ignore
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('wpTitle', {
                  label: 'Webpart címe'
                }),
                PropertyPaneChoiceGroup('position', { 
                  label: "Kártya poziciója",
                  options: [
                    {key: 'topCenter', text: 'Fent'},
                    {key: 'bottomCenter', text: 'Lent'},
                    {key: 'leftCenter', text: 'Bal'},
                    {key: 'rightCenter', text: 'Jobb'}
                  ]        
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
