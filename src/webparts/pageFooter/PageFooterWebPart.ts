import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  PropertyFieldFilePicker,
  IFilePickerResult
} from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';
import * as strings from 'PageFooterWebPartStrings';

import PageFooter from './components/PageFooter';
import { IPageFooterProps } from './components/IPageFooterProps';

export interface IPageFooterWebPartProps {
  footerListName: string;
  linkedInUrl: string;
  youTubeUrl: string;
  logoUrl: string;
  logo: IFilePickerResult;
}

export default class PageFooterWebPart extends BaseClientSideWebPart<IPageFooterWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPageFooterProps> = React.createElement(
      PageFooter,
      {
        footerListName: this.properties.footerListName || 'FooterLinks',
        linkedInUrl: this.properties.linkedInUrl,
        youTubeUrl: this.properties.youTubeUrl,
        logoUrl: this.properties.logoUrl,
        context: this.context
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
          header: { description: strings.PropertyPaneHeader },
          groups: [
            {
              groupName: strings.GroupNameGeneral,
              groupFields: [
                PropertyPaneTextField('footerListName', {
                  label: strings.FieldFooterListName
                }),
                PropertyFieldFilePicker('logo', {
                  context: this.context as any,
                  filePickerResult: this.properties.logo,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (e: IFilePickerResult) => {
                    this.properties.logo = e;
                    this.properties.logoUrl = e.fileAbsoluteUrl;
                    this.render();
                  },
                  onChanged: (e: IFilePickerResult) => {
                    this.properties.logo = e;
                  },
                  key: 'logoImagePicker',
                  label: strings.FieldLogoImage
                })
              ]
            },
            {
              groupName: strings.GroupNameSocial,
              groupFields: [
                PropertyPaneTextField('linkedInUrl', {
                  label: strings.FieldLinkedInUrl
                }),
                PropertyPaneTextField('youTubeUrl', {
                  label: strings.FieldYouTubeUrl
                })
              ]
            }
          ]
        }
      ]
    };
  }
}