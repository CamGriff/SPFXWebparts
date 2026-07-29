import * as React from 'react';
import {
  TextField,
  PrimaryButton,
  MessageBar,
  MessageBarType,
  Stack,
  IconButton,
  Spinner,
  SpinnerSize,
  Text,
  useTheme,
  mergeStyles
} from '@fluentui/react';
import { SPFI } from '@pnp/sp';

import { IConfigItem } from '../../../../models/IConfigItem';
import { getConfigs, saveConfig, deleteConfig } from '../../../../services/ConfigService';

export interface IConfigPanelProps {
  sp: SPFI;
}

const DEFAULT_STALE_DAYS = 7;

interface IConfigFormState {
  siteUrl: string;
  language: string;
  staleDays: string;
  translatorName: string;
  translatorEmail: string;
}

const EMPTY_FORM: IConfigFormState = {
  siteUrl: '',
  language: 'fr-fr',
  staleDays: String(DEFAULT_STALE_DAYS),
  translatorName: '',
  translatorEmail: ''
};

type TextFieldChangeEvent = React.FormEvent<HTMLInputElement | HTMLTextAreaElement>;

function ConfigPanel(props: IConfigPanelProps): JSX.Element {
  const theme = useTheme();
  const [configs, setConfigs] = React.useState<IConfigItem[]>([]);
  const [form, setForm] = React.useState<IConfigFormState>(EMPTY_FORM);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [isSaving, setIsSaving] = React.useState<boolean>(false);
  const [errorMessage, setErrorMessage] = React.useState<string | undefined>(undefined);

  const loadConfigs = React.useCallback(async (): Promise<void> => {
    setIsLoading(true);
    setErrorMessage(undefined);

    try {
      const results = await getConfigs(props.sp);
      setConfigs(results);
    } catch (error) {
      console.error('Failed to load registered sites.', error);
      setErrorMessage('Failed to load registered sites.');
    } finally {
      setIsLoading(false);
    }
  }, [props.sp]);

  React.useEffect(() => {
    void loadConfigs();
  }, [loadConfigs]);

  const updateField = (field: keyof IConfigFormState): ((event?: TextFieldChangeEvent, newValue?: string) => void) => {
    return (_event?: TextFieldChangeEvent, newValue?: string): void => {
      setForm((current) => ({ ...current, [field]: newValue ?? '' }));
    };
  };

  const handleSave = async (): Promise<void> => {
    if (!form.siteUrl || !form.language) {
      setErrorMessage('Site URL and language are required.');
      return;
    }

    const parsedStaleDays = parseInt(form.staleDays, 10);

    setIsSaving(true);
    setErrorMessage(undefined);

    try {
      const config: IConfigItem = {
        Title: form.siteUrl,
        SiteUrl: form.siteUrl,
        Language: form.language.toLowerCase(),
        StaleDays: isNaN(parsedStaleDays) ? DEFAULT_STALE_DAYS : parsedStaleDays,
        TranslatorName: form.translatorName,
        TranslatorEmail: form.translatorEmail,
        IsActive: true
      };

      await saveConfig(props.sp, config);
      setForm(EMPTY_FORM);
      await loadConfigs();
    } catch (error) {
      console.error('Failed to save site configuration.', error);
      setErrorMessage('Failed to save site configuration.');
    } finally {
      setIsSaving(false);
    }
  };

  const handleDelete = async (id: number | undefined): Promise<void> => {
    if (id === undefined) {
      return;
    }

    setErrorMessage(undefined);

    try {
      await deleteConfig(props.sp, id);
      await loadConfigs();
    } catch (error) {
      console.error('Failed to delete site configuration.', error);
      setErrorMessage('Failed to delete site configuration.');
    }
  };

  const panelClass = mergeStyles({
    padding: 16,
    border: `1px solid ${theme.palette.neutralLight}`,
    borderRadius: theme.effects.roundedCorner4,
    backgroundColor: theme.semanticColors.bodyBackground
  });

  const rowClass = mergeStyles({
    padding: '6px 0',
    borderBottom: `1px solid ${theme.palette.neutralLighter}`
  });

  return (
    <div className={panelClass}>
      <Stack tokens={{ childrenGap: 16 }}>
        <Text variant="large" styles={{ root: { fontWeight: 600, color: theme.semanticColors.bodyText } }}>
          Registered Sites
        </Text>

        {errorMessage && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setErrorMessage(undefined)}>
            {errorMessage}
          </MessageBar>
        )}

        <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
          <TextField
            label="Site URL"
            value={form.siteUrl}
            onChange={updateField('siteUrl')}
            placeholder="https://contoso.sharepoint.com/sites/example"
            styles={{ root: { minWidth: 260 } }}
          />
          <TextField
            label="Language"
            value={form.language}
            onChange={updateField('language')}
            placeholder="fr-fr"
            styles={{ root: { minWidth: 100 } }}
          />
          <TextField
            label="Stale Days"
            value={form.staleDays}
            onChange={updateField('staleDays')}
            type="number"
            styles={{ root: { minWidth: 90 } }}
          />
          <TextField
            label="Translator Name"
            value={form.translatorName}
            onChange={updateField('translatorName')}
            styles={{ root: { minWidth: 160 } }}
          />
          <TextField
            label="Translator Email"
            value={form.translatorEmail}
            onChange={updateField('translatorEmail')}
            styles={{ root: { minWidth: 200 } }}
          />
        </Stack>

        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text="Add Site"
            onClick={() => {
              void handleSave();
            }}
            disabled={isSaving}
          />
          {isSaving && <Spinner size={SpinnerSize.small} label="Saving..." labelPosition="right" />}
        </Stack>

        {isLoading ? (
          <Spinner size={SpinnerSize.medium} label="Loading registered sites..." />
        ) : configs.length === 0 ? (
          <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary } }}>
            No sites registered yet.
          </Text>
        ) : (
          <Stack tokens={{ childrenGap: 0 }}>
            {configs.map((config) => (
              <Stack
                key={`${config.SiteUrl}|${config.Language}`}
                horizontal
                verticalAlign="center"
                tokens={{ childrenGap: 12 }}
                className={rowClass}
              >
                <Stack.Item grow>
                  <Text variant="small" styles={{ root: { color: theme.semanticColors.bodyText } }}>
                    {config.SiteUrl}
                  </Text>
                </Stack.Item>
                <Stack.Item>
                  <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                    {config.Language}
                  </Text>
                </Stack.Item>
                <Stack.Item>
                  <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                    {`Stale > ${config.StaleDays}d`}
                  </Text>
                </Stack.Item>
                <Stack.Item>
                  <Text variant="small" styles={{ root: { color: theme.semanticColors.bodyText } }}>
                    {config.TranslatorName || '—'}
                  </Text>
                </Stack.Item>
                <Stack.Item>
                  <IconButton
                    iconProps={{ iconName: 'Delete' }}
                    title="Remove"
                    ariaLabel="Remove"
                    styles={{ root: { color: theme.palette.neutralSecondary } }}
                    onClick={() => {
                      void handleDelete(config.Id);
                    }}
                  />
                </Stack.Item>
              </Stack>
            ))}
          </Stack>
        )}
      </Stack>
    </div>
  );
}

export default ConfigPanel;
