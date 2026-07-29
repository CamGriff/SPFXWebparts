import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';

import { provisionLists } from '../../services/ListProvisioner';

const LOG_SOURCE = 'MultilingualGovernanceCustomiserApplicationCustomizer';

export interface IMultilingualGovernanceCustomiserApplicationCustomizerProperties {
}

export default class MultilingualGovernanceCustomiserApplicationCustomizer
  extends BaseApplicationCustomizer<IMultilingualGovernanceCustomiserApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    const sp = spfi().using(SPFx(this.context));

    void provisionLists(sp).catch((error: Error) => {
      Log.error(LOG_SOURCE, error);
    });

    return Promise.resolve();
  }
}
