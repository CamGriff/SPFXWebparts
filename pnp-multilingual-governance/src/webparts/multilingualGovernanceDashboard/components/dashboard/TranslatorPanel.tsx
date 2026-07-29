import * as React from 'react';
import { Stack, PrimaryButton, Spinner, SpinnerSize, Text, useTheme, mergeStyles } from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import { IDriftItem } from '../../../../models/IDriftItem';
import { DriftStatus } from '../../../../models/DriftStatus';
import { sendNudgeEmail, markNudgeSent } from '../../../../services/NotificationService';

export interface ITranslatorPanelProps {
  allItems: IDriftItem[];
  sp: SPFI;
  graphClient: MSGraphClientV3;
}

interface ITranslatorGroup {
  key: string;
  translatorEmail: string;
  translatorName: string;
  siteUrl: string;
  staleCount: number;
  missingCount: number;
  abandonedCount: number;
  pages: IDriftItem[];
}

const OUTSTANDING_STATUSES = new Set<string>([DriftStatus.Stale, DriftStatus.Missing, DriftStatus.Abandoned]);

function buildGroups(items: IDriftItem[]): ITranslatorGroup[] {
  const map = new Map<string, ITranslatorGroup>();

  for (const item of items) {
    if (!item.TranslatorEmail || !OUTSTANDING_STATUSES.has(item.DriftStatus)) {
      continue;
    }

    const key = `${item.TranslatorEmail.toLowerCase()}|${item.SiteUrl.toLowerCase()}|${item.TranslationLanguage.toLowerCase()}`;

    if (!map.has(key)) {
      map.set(key, {
        key,
        translatorEmail: item.TranslatorEmail,
        translatorName: item.TranslatorName ?? '',
        siteUrl: item.SiteUrl,
        staleCount: 0,
        missingCount: 0,
        abandonedCount: 0,
        pages: []
      });
    }

    const group = map.get(key)!;
    group.pages.push(item);

    if (item.DriftStatus === DriftStatus.Stale) group.staleCount++;
    else if (item.DriftStatus === DriftStatus.Missing) group.missingCount++;
    else if (item.DriftStatus === DriftStatus.Abandoned) group.abandonedCount++;
  }

  return Array.from(map.values());
}

interface ITranslatorCardProps {
  group: ITranslatorGroup;
  sp: SPFI;
  graphClient: MSGraphClientV3;
}

function TranslatorCard(props: ITranslatorCardProps): JSX.Element {
  const { group } = props;
  const theme = useTheme();
  const [isSending, setIsSending] = React.useState<boolean>(false);
  const [isSent, setIsSent] = React.useState<boolean>(false);
  const [showSuccess, setShowSuccess] = React.useState<boolean>(false);
  const [errorMessage, setErrorMessage] = React.useState<string | undefined>(undefined);

  const handleNudge = async (): Promise<void> => {
    setIsSending(true);
    setErrorMessage(undefined);

    try {
      await sendNudgeEmail(
        props.graphClient,
        group.translatorEmail,
        group.translatorName,
        group.pages,
        group.siteUrl
      );
      await markNudgeSent(
        props.sp,
        group.pages.map((p) => p.PageGuid),
        group.pages[0].TranslationLanguage
      );
      setIsSent(true);
      setShowSuccess(true);
      setTimeout(() => setShowSuccess(false), 3000);
    } catch {
      setErrorMessage('Failed to send nudge');
    } finally {
      setIsSending(false);
    }
  };

  const total = group.staleCount + group.missingCount + group.abandonedCount;

  const cardClass = mergeStyles({
    border: `1px solid ${theme.palette.neutralLight}`,
    borderRadius: theme.effects.roundedCorner4,
    padding: '16px',
    backgroundColor: theme.semanticColors.bodyBackground,
    minWidth: 260,
    maxWidth: 320,
    boxShadow: theme.effects.elevation4
  });

  const countItems = [
    { label: 'Stale', count: group.staleCount },
    { label: 'Missing', count: group.missingCount },
    { label: 'Abandoned', count: group.abandonedCount },
    { label: 'Total outstanding', count: total, isTotal: true }
  ];

  return (
    <div className={cardClass}>
      <Stack tokens={{ childrenGap: 8 }}>
        <Text
          variant="mediumPlus"
          styles={{ root: { fontWeight: 700, color: theme.semanticColors.bodyText } }}
        >
          {group.translatorName || group.translatorEmail}
        </Text>
        <Text
          variant="small"
          styles={{ root: { color: theme.palette.neutralSecondary, wordBreak: 'break-all' } }}
        >
          {group.siteUrl}
        </Text>

        <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
          {countItems.map(({ label, count, isTotal }) => (
            <Stack key={label}>
              <Text variant="xSmall" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                {label}
              </Text>
              <Text
                variant="mediumPlus"
                styles={{
                  root: {
                    fontWeight: 600,
                    color: isTotal && count > 0 ? theme.palette.redDark : theme.palette.neutralPrimary
                  }
                }}
              >
                {count}
              </Text>
            </Stack>
          ))}
        </Stack>

        {errorMessage && (
          <Text variant="small" styles={{ root: { color: theme.semanticColors.errorText } }}>
            {errorMessage}
          </Text>
        )}

        {isSending ? (
          <Spinner size={SpinnerSize.small} label="Sending..." labelPosition="right" />
        ) : showSuccess ? (
          <Text variant="small" styles={{ root: { color: theme.palette.green, fontWeight: 600 } }}>
            Nudge sent!
          </Text>
        ) : (
          <PrimaryButton
            text={isSent ? 'Sent' : `Nudge ${group.translatorName || group.translatorEmail}`}
            disabled={isSent}
            onClick={() => {
              void handleNudge();
            }}
          />
        )}
      </Stack>
    </div>
  );
}

function TranslatorPanel(props: ITranslatorPanelProps): JSX.Element | null {
  const theme = useTheme();
  const groups = buildGroups(props.allItems);

  if (groups.length === 0) {
    return null;
  }

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <Text variant="medium" styles={{ root: { fontWeight: 600, color: theme.semanticColors.bodyText } }}>
        Translator outstanding work
      </Text>
      <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
        {groups.map((group) => (
          <TranslatorCard key={group.key} group={group} sp={props.sp} graphClient={props.graphClient} />
        ))}
      </Stack>
    </Stack>
  );
}

export default TranslatorPanel;
