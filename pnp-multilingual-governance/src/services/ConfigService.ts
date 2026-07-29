import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { IList } from '@pnp/sp/lists';

import { IConfigItem } from '../models/IConfigItem';

const GOVERNANCE_CONFIG_LIST_TITLE = 'GovernanceConfig';
const MAX_ITEMS_PER_REQUEST = 5000;

interface IExistingConfigRow {
  Id: number;
}

export async function getConfigs(sp: SPFI): Promise<IConfigItem[]> {
  const configs: IConfigItem[] = await sp.web.lists
    .getByTitle(GOVERNANCE_CONFIG_LIST_TITLE)
    .items.select('Id', 'Title', 'SiteUrl', 'Language', 'StaleDays', 'TranslatorName', 'TranslatorEmail', 'IsActive')
    .filter('IsActive eq 1')
    .top(MAX_ITEMS_PER_REQUEST)();

  return configs;
}

export async function saveConfig(sp: SPFI, config: IConfigItem): Promise<void> {
  const list = sp.web.lists.getByTitle(GOVERNANCE_CONFIG_LIST_TITLE);
  const existing = await findExistingConfig(list, config.SiteUrl, config.Language);

  const fields = {
    Title: config.Title,
    SiteUrl: config.SiteUrl,
    Language: config.Language,
    StaleDays: config.StaleDays,
    TranslatorName: config.TranslatorName,
    TranslatorEmail: config.TranslatorEmail,
    IsActive: config.IsActive
  };

  if (existing) {
    await list.items.getById(existing.Id).update(fields);
  } else {
    await list.items.add(fields);
  }
}

export async function deleteConfig(sp: SPFI, id: number): Promise<void> {
  await sp.web.lists.getByTitle(GOVERNANCE_CONFIG_LIST_TITLE).items.getById(id).delete();
}

async function findExistingConfig(list: IList, siteUrl: string, language: string): Promise<IExistingConfigRow | undefined> {
  const filter = `SiteUrl eq '${escapeODataString(siteUrl)}' and Language eq '${escapeODataString(language)}'`;
  const matches: IExistingConfigRow[] = await list.items.select('Id').filter(filter).top(1)();

  return matches[0];
}

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}
