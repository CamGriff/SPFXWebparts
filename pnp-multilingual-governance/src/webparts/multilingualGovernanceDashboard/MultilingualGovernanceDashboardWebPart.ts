import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Log, Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneSlider } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { SPFI, spfi } from '@pnp/sp';
import { SPFx } from '@pnp/sp/behaviors/spfx';
import '@pnp/sp/webs';

import { ThemeProvider } from '@fluentui/react';

import * as strings from 'MultilingualGovernanceDashboardWebPartStrings';
import MultilingualGovernanceDashboard from './components/MultilingualGovernanceDashboard';
import { IMultilingualGovernanceDashboardProps } from './components/IMultilingualGovernanceDashboardProps';
import { provisionLists } from '../../services/ListProvisioner';

const LOG_SOURCE = 'MultilingualGovernanceDashboardWebPart';

export interface IMultilingualGovernanceDashboardWebPartProps {
  listName: string;
  staleDaysThreshold: number;
}

export default class MultilingualGovernanceDashboardWebPart extends BaseClientSideWebPart<IMultilingualGovernanceDashboardWebPartProps> {
  private sp!: SPFI;
  private graphClient!: MSGraphClientV3;

  protected async onInit(): Promise<void> {
    this.sp = spfi().using(SPFx(this.context));
    this.graphClient = await this.context.msGraphClientFactory.getClient('3');

    // Redundant provisioning trigger: the Application Customiser may not have run yet on
    // this site (feature not activated, or it ran under a visitor without Manage Lists
    // rights and failed silently). Whoever adds this web part is far more likely to have
    // owner-level permissions, so retry provisioning here too.
    try {
      await provisionLists(this.sp);
    } catch (error) {
      Log.error(LOG_SOURCE, error as Error);
    }
  }

  public render(): void {
    const dashboard = React.createElement(MultilingualGovernanceDashboard, {
      context: this.context,
      sp: this.sp,
      graphClient: this.graphClient
    } as IMultilingualGovernanceDashboardProps);

    const element = React.createElement(ThemeProvider, null, dashboard);

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
                PropertyPaneTextField('listName', {
                  label: 'Translation Drift list name',
                  value: this.properties.listName
                }),
                PropertyPaneSlider('staleDaysThreshold', {
                  label: 'Default stale threshold (days)',
                  min: 1,
                  max: 90,
                  value: this.properties.staleDaysThreshold
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
