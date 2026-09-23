import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'EnterprisePropertiesWpWebPartStrings';
import EnterprisePropertiesWp from './components/EnterprisePropertiesWp';
import { IEnterprisePropertiesWpProps } from './components/IEnterprisePropertiesWpProps';

export interface IEnterprisePropertiesWpWebPartProps {
  title: string;
}

export default class EnterprisePropertiesWpWebPart
  extends BaseClientSideWebPart<IEnterprisePropertiesWpWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IEnterprisePropertiesWpProps> = React.createElement(
      EnterprisePropertiesWp,
      {
        title: this.properties.title || 'Enterprise Priorities',
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    if (!this.properties.title) {
      this.properties.title = 'Enterprise Priorities';
    }

    return Promise.resolve();
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