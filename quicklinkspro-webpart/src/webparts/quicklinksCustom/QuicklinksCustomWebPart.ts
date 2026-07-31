import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import {
  PropertyFieldCollectionData,
  CustomCollectionFieldType
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import * as strings from 'QuicklinksCustomWebPartStrings';

import QuicklinksCustom from './components/QuicklinksCustom';
import { IQuicklinksCustomProps, IQuickLinkTile } from './components/IQuicklinksCustomProps';

export interface IQuicklinksCustomWebPartProps {
  tiles: IQuickLinkTile[];
}

export default class QuicklinksCustomWebPart extends BaseClientSideWebPart<IQuicklinksCustomWebPartProps> {

  private _isDarkTheme: boolean = false;

  public render(): void {
    const element: React.ReactElement<IQuicklinksCustomProps> = React.createElement(
      QuicklinksCustom,
      {
        tiles: this.properties.tiles || [],
        isDarkTheme: this._isDarkTheme,
        displayMode: this.displayMode
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
              groupName: strings.GroupNameTiles,
              groupFields: [
                PropertyFieldCollectionData('tiles', {
                  key: 'tiles',
                  label: strings.FieldTiles,
                  panelHeader: strings.FieldTiles,
                  manageBtnLabel: 'Manage Tiles',
                  value: this.properties.tiles || [],
                  fields: [
                    {
                      id: 'title',
                      title: 'Title',
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: 'description',
                      title: 'Description',
                      type: CustomCollectionFieldType.string,
                      required: false
                    },
                    {
                      id: 'url',
                      title: 'URL',
                      type: CustomCollectionFieldType.url,
                      required: true
                    },
                    {
                      id: 'icon',
                      title: 'Icon (Fluent UI)',
                      type: CustomCollectionFieldType.string,
                      required: false
                    }
                  ],
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}