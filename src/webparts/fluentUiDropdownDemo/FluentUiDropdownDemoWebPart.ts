import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'FluentUiDropdownDemoWebPartStrings';
import FluentUiDropdownDemo from './components/FluentUiDropdownDemo';
import { IFluentUiDropdownDemoProps } from './components/IFluentUiDropdownDemoProps';

export interface IFluentUiDropdownDemoWebPartProps {
  description: string;
}

export default class FluentUiDropdownDemoWebPart extends BaseClientSideWebPart<IFluentUiDropdownDemoWebPartProps> {

  public async render(): Promise<void> {
    const element: React.ReactElement<IFluentUiDropdownDemoProps> = React.createElement(
      FluentUiDropdownDemo,
      {
        description: this.properties.description,
        webURL:this.context.pageContext.web.absoluteUrl,
        singleValueOptions:await getChoiceFields(this.context.pageContext.web.absoluteUrl,'SingleValueDropdown'),
        multiValueOptions:await getChoiceFields(this.context.pageContext.web.absoluteUrl,'MultiValueDropdown')
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
export const getChoiceFields = async (webURL,field) => {
  let resultarr = [];
  await fetch(`${webURL}/_api/web/lists/GetByTitle('FluentUIDropdown')/fields?$filter=EntityPropertyName eq '${field}'`, {
      method: 'GET',
      mode: 'cors',
      credentials: 'same-origin',
      headers: new Headers({
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'Access-Control-Allow-Origin': '*',
          'Cache-Control': 'no-cache',
          'pragma': 'no-cache',
      }),
  }).then(async (response) => await response.json())
      .then(async (data) => {
          for (var i = 0; i < data.value[0].Choices.length; i++) {
              
              await resultarr.push({
                  key:data.value[0].Choices[i],
                  text:data.value[0].Choices[i]
                  
            });
          
      }
      });
  return await resultarr;
};