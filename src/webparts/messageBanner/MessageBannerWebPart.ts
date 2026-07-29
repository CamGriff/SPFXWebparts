import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from '@pnp/sp';

import MessageBanner from './components/MessageBanner';
import { IMessageBannerProps } from './components/IMessageBannerProps';

export default class MessageBannerWebPart extends BaseClientSideWebPart<object> {

public render(): void {
  this.domElement.style.display = 'none'; // hide by default

  const isFrench = this.context.pageContext.cultureInfo.currentUICultureName
    .toLowerCase()
    .startsWith('fr');

  const sp = spfi().using(SPFx(this.context));

  const element: React.ReactElement<IMessageBannerProps> = React.createElement(
    MessageBanner,
    {
      sp,
      isFrench,
      onNoItems: () => {
        this.domElement.style.display = 'none';
      },
      onHasItems: () => {
        this.domElement.style.display = 'block';
      }
    }
  );

  ReactDom.render(element, this.domElement);
}

  protected onDispose(): void {
  ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get isRenderAsync(): boolean {
    return true;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}