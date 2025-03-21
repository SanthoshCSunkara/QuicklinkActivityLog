import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import * as strings from 'QuicklinkActivityLogWebPartStrings';
import QuicklinkActivityLog from './components/QuicklinkActivityLog';
import { IQuicklinkActivityLogProps } from './components/IQuicklinkActivityLogProps';

export interface IQuicklinkActivityLogWebPartProps {
  webPartTitle: string;
}

export default class QuicklinkActivityLogWebPart extends BaseClientSideWebPart<IQuicklinkActivityLogWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IQuicklinkActivityLogProps> = React.createElement(
      QuicklinkActivityLog,
      {
        context: this.context,
        webPartTitle: this.properties.webPartTitle,
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
                PropertyPaneTextField('webPartTitle', {
                  label: 'Web Part Title'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}