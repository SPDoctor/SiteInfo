import { Version } from '@microsoft/sp-core-library';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox

} from '@microsoft/sp-webpart-base';

import SiteInfo, {ISiteInfoProps }  from './components/SiteInfo';
//import { IHelloWorldWebPartProps } from './ISiteInfoWebPartProps';
import { ISiteInfoWebPartProps } from './ISiteInfoWebPartProps';

export default class SiteInfoWebPart extends BaseClientSideWebPart<ISiteInfoWebPartProps> {


  public render(): void {
    const element: React.ReactElement<ISiteInfoWebPartProps> = React.createElement(SiteInfo, {
      description: this.properties.description,
      showLists: this.properties.showLists,
      showUser: this.properties.showUser,
      self: this
    });

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
            description: "SiteInfo Settings"
          },
          groups: [
            {
              groupName: "Properties",
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Description Field",
                  placeholder: "enter a description"
                }),
                PropertyPaneCheckbox('showLists', {
                  checked: false,
                  text: "Show Lists"
                }),
                PropertyPaneCheckbox('showUser', {
                  checked: false,
                  text: "Show User"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
