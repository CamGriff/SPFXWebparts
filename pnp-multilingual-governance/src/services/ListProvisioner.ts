import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import { IList } from '@pnp/sp/lists';
import { DriftStatus } from '../models/DriftStatus';

const TRANSLATION_DRIFT_LIST_TITLE = 'TranslationDrift';
const GOVERNANCE_CONFIG_LIST_TITLE = 'GovernanceConfig';
const GENERIC_LIST_TEMPLATE = 100;
const DRIFT_STATUS_CHOICES: string[] = [
  DriftStatus.InSync,
  DriftStatus.Stale,
  DriftStatus.Missing,
  DriftStatus.Orphaned,
  DriftStatus.Abandoned
];

interface IFieldDefinition {
  internalName: string;
  add: (list: IList) => Promise<unknown>;
}

export async function provisionLists(sp: SPFI): Promise<void> {
  const driftList = await ensureList(
    sp,
    TRANSLATION_DRIFT_LIST_TITLE,
    'Tracks translation drift between source pages and their translations'
  );
  await ensureFields(driftList, getTranslationDriftFieldDefinitions());

  const configList = await ensureList(
    sp,
    GOVERNANCE_CONFIG_LIST_TITLE,
    'Sites and translators registered for multilingual governance monitoring'
  );
  await ensureFields(configList, getGovernanceConfigFieldDefinitions());
}

async function ensureList(sp: SPFI, title: string, description: string): Promise<IList> {
  const ensureResult = await sp.web.lists.ensure(title, description, GENERIC_LIST_TEMPLATE, true);
  return ensureResult.list;
}

async function ensureFields(list: IList, definitions: IFieldDefinition[]): Promise<void> {
  const existingFields = await list.fields.select('InternalName')();
  const existingNames = new Set(existingFields.map((field) => field.InternalName));

  for (const definition of definitions) {
    if (!existingNames.has(definition.internalName)) {
      await definition.add(list);
    }
  }
}

function getTranslationDriftFieldDefinitions(): IFieldDefinition[] {
  return [
    { internalName: 'DefaultPageTitle', add: (list) => list.fields.addText('DefaultPageTitle') },
    { internalName: 'DefaultPageUrl', add: (list) => list.fields.addUrl('DefaultPageUrl') },
    { internalName: 'DefaultPageModified', add: (list) => list.fields.addDateTime('DefaultPageModified') },
    { internalName: 'TranslationLanguage', add: (list) => list.fields.addText('TranslationLanguage') },
    { internalName: 'TranslationPageUrl', add: (list) => list.fields.addUrl('TranslationPageUrl') },
    { internalName: 'TranslationModified', add: (list) => list.fields.addDateTime('TranslationModified') },
    { internalName: 'DaysDrift', add: (list) => list.fields.addNumber('DaysDrift') },
    {
      internalName: 'DriftStatus',
      add: (list) => list.fields.addChoice('DriftStatus', { Choices: DRIFT_STATUS_CHOICES })
    },
    { internalName: 'TranslatorName', add: (list) => list.fields.addText('TranslatorName') },
    { internalName: 'TranslatorEmail', add: (list) => list.fields.addText('TranslatorEmail') },
    { internalName: 'SiteUrl', add: (list) => list.fields.addText('SiteUrl') },
    { internalName: 'LastChecked', add: (list) => list.fields.addDateTime('LastChecked') },
    { internalName: 'PageGuid', add: (list) => list.fields.addText('PageGuid') },
    { internalName: 'NudgeSent', add: (list) => list.fields.addBoolean('NudgeSent') },
    { internalName: 'NudgeDate', add: (list) => list.fields.addDateTime('NudgeDate') }
  ];
}

function getGovernanceConfigFieldDefinitions(): IFieldDefinition[] {
  return [
    { internalName: 'SiteUrl', add: (list) => list.fields.addText('SiteUrl') },
    { internalName: 'Language', add: (list) => list.fields.addText('Language') },
    { internalName: 'StaleDays', add: (list) => list.fields.addNumber('StaleDays') },
    { internalName: 'TranslatorName', add: (list) => list.fields.addText('TranslatorName') },
    { internalName: 'TranslatorEmail', add: (list) => list.fields.addText('TranslatorEmail') },
    { internalName: 'IsActive', add: (list) => list.fields.addBoolean('IsActive') }
  ];
}
