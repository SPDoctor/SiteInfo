import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField,
  PropertyPaneCheckbox
} from '@microsoft/sp-client-preview';

import SiteInfo, { ISiteInfoProps } from './components/SiteInfo';
import { ISiteInfoWebPartProps } from './ISiteInfoWebPartProps';

export default class SiteInfoWebPart extends BaseClientSideWebPart<ISiteInfoWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    const element: React.ReactElement<ISiteInfoProps> = React.createElement(SiteInfo, {
      description: this.properties.description,
      showLists: this.properties.showLists,
      showUser: this.properties.showUser,
      self: this
    });

    ReactDom.render(element, this.domElement);
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
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
                  isChecked: false,
                  text: "Show Lists"
                }),
                PropertyPaneCheckbox('showUser', {
                  isChecked: false,
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
