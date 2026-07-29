import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { IBannerItem } from '../models/IBannerItem';

const LIST_NAME = 'MessageBanner';

export class BannerService {
  constructor(private sp: SPFI) {}

  public async getActiveItems(): Promise<IBannerItem[]> {
    const today = new Date().toISOString();

    const items = await this.sp.web.lists
      .getByTitle(LIST_NAME)
      .items.filter(
        `MessageStatus eq 'Active' and (ExpirationDate ge datetime'${today}' or ExpirationDate eq null)`
      )
      .select(
        'Id',
        'Title',
        'TitleFR',
        'Description',
        'DescriptionFR',
        'SeeMore/Url',
        'SeeMore/Description',
        'SeeMoreFR/Url',
        'SeeMoreFR/Description',
        'MessageLevel',
        'MessageStatus',
        'ExpirationDate'
      )
      .orderBy('MessageLevel', false)();

    return items.map((item: Record<string, any>) => ({
      Id: item.Id,
      Title: item.Title,
      TitleFR: item.TitleFR,
      Description: item.Description,
      DescriptionFR: item.DescriptionFR,
      SeeMoreUrl: item.SeeMore?.Url ?? '',
      SeeMoreDescription: item.SeeMore?.Description ?? '',
      SeeMoreFRUrl: item.SeeMoreFR?.Url ?? '',
      SeeMoreFRDescription: item.SeeMoreFR?.Description ?? '',
      MessageLevel: item.MessageLevel,
      MessageStatus: item.MessageStatus,
      ExpirationDate: item.ExpirationDate,
    }));
  }
}