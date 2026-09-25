import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import { MegaMenu, IMegaMenuProps } from './components/MegaMenu';

const LOG_SOURCE: string = 'CustomMegamenuApplicationCustomizer';

export interface ICustomMegamenuApplicationCustomizerProperties {
  // Reserved for client-side component properties passed via the custom
  // action registration, if any are needed later.
}

export default class CustomMegamenuApplicationCustomizer
  extends BaseApplicationCustomizer<ICustomMegamenuApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized CustomMegamenuApplicationCustomizer');

    // Placeholders may not be available immediately; this event fires
    // whenever they change, so wiring the render through it is safer than
    // assuming the Top placeholder exists at onInit time.
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    this._renderPlaceHolders();

    return Promise.resolve();
  }

  private _renderPlaceHolders(): void {
    // Only create the Top placeholder content once. changedEvent can fire
    // more than once; without this guard we'd stack duplicate menus.
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );

      // The Top placeholder isn't guaranteed to exist on every page/host;
      // bail cleanly rather than throwing if it isn't available.
      if (!this._topPlaceholder) {
        Log.info(LOG_SOURCE, 'Top placeholder not available on this page.');
        return;
      }
    }

    if (this._topPlaceholder.domElement) {
      const element: React.ReactElement<IMegaMenuProps> = React.createElement(
        MegaMenu,
        {
          // Hardcoded for this first proof-of-life render. Swapped for a
          // SharePoint list read once the shell is confirmed mounting.
          context: this.context
        }
      );
      ReactDOM.render(element, this._topPlaceholder.domElement);
    }
  }

  private _onDispose = (): void => {
    if (this._topPlaceholder && this._topPlaceholder.domElement) {
      ReactDOM.unmountComponentAtNode(this._topPlaceholder.domElement);
    }
    Log.info(LOG_SOURCE, 'Disposed custom mega menu.');
  }
}