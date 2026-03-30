import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import {
  PropertyFieldFilePicker,
  IFilePickerResult
} from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';
import * as strings from 'PageHeaderWebPartStrings';

import PageHeader from './components/PageHeader';
import { IPageHeaderProps } from './components/IPageHeaderProps';

export interface IPageHeaderWebPartProps {
  siteTitle: string;
  searchBoxPlaceholder: string;
  searchPageUrl: string;
  backgroundImageUrl: string;
  backgroundImage: IFilePickerResult;
  seasonalEnabled: boolean;
  seasonalLabel: string;
  seasonalUrl: string;
}

export default class PageHeaderWebPart extends BaseClientSideWebPart<IPageHeaderWebPartProps> {

  private _isDarkTheme: boolean = false;

  public render(): void {
    const element: React.ReactElement<IPageHeaderProps> = React.createElement(
      PageHeader,
      {
        siteTitle: this.properties.siteTitle,
        searchBoxPlaceholder: this.properties.searchBoxPlaceholder,
        searchPageUrl: this.properties.searchPageUrl,
        backgroundImageUrl: this.properties.backgroundImageUrl,
        isDarkTheme: this._isDarkTheme,
        seasonalEnabled: this.properties.seasonalEnabled,
        seasonalLabel: this.properties.seasonalLabel,
        seasonalUrl: this.properties.seasonalUrl,
        context: this.context
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
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
                PropertyPaneTextField('siteTitle', {
                  label: strings.FieldSiteTitle
                }),
                PropertyPaneTextField('searchBoxPlaceholder', {
                  label: strings.FieldSearchPlaceholder
                }),
                PropertyPaneTextField('searchPageUrl', {
                  label: strings.FieldSearchPageUrl
                }),
                PropertyFieldFilePicker('backgroundImage', {
                  context: this.context as any,
                  filePickerResult: this.properties.backgroundImage,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (e: IFilePickerResult) => {
                    this.properties.backgroundImage = e;
                    this.properties.backgroundImageUrl = e.fileAbsoluteUrl;
                    this.render();
                  },
                  onChanged: (e: IFilePickerResult) => {
                    this.properties.backgroundImage = e;
                  },
                  key: 'backgroundImagePicker',
                  label: strings.FieldBackgroundImage
                })
              ]
            },
            {
              groupName: strings.GroupNameSeasonal,
              groupFields: [
                PropertyPaneToggle('seasonalEnabled', {
                  label: strings.FieldSeasonalEnabled
                }),
                PropertyPaneTextField('seasonalLabel', {
                  label: strings.FieldSeasonalLabel,
                  disabled: !this.properties.seasonalEnabled
                }),
                PropertyPaneTextField('seasonalUrl', {
                  label: strings.FieldSeasonalUrl,
                  disabled: !this.properties.seasonalEnabled
                })
              ]
            }
          ]
        }
      ]
    };
  }
}