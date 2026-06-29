import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';
import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IMultilingualGovernanceDashboardProps {
  context: WebPartContext;
  sp: SPFI;
  graphClient: MSGraphClientV3;
}
