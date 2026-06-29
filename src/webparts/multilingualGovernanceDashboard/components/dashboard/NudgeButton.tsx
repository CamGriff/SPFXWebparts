import * as React from 'react';
import { DefaultButton, Spinner, SpinnerSize } from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import { IDriftItem } from '../../../../models/IDriftItem';
import { DriftStatus } from '../../../../models/DriftStatus';
import { sendNudgeEmail, markNudgeSent } from '../../../../services/NotificationService';

export interface INudgeButtonProps {
  item: IDriftItem;
  allItems: IDriftItem[];
  sp: SPFI;
  graphClient: MSGraphClientV3;
  onNudgeSent?: (item: IDriftItem) => void;
}

const NUDGEABLE_STATUSES = new Set<DriftStatus>([DriftStatus.Stale, DriftStatus.Missing]);

function NudgeButton(props: INudgeButtonProps): JSX.Element {
  const [isSending, setIsSending] = React.useState<boolean>(false);
  const [isSent, setIsSent] = React.useState<boolean>(props.item.NudgeSent);
  const [showSuccess, setShowSuccess] = React.useState<boolean>(false);
  const [errorMessage, setErrorMessage] = React.useState<string | undefined>(undefined);

  const handleNudge = async (): Promise<void> => {
    const { TranslatorEmail, TranslatorName, SiteUrl, TranslationLanguage } = props.item;
    if (!TranslatorEmail) {
      return;
    }

    const pagesToNudge = props.allItems.filter(
      (p) =>
        p.SiteUrl === SiteUrl &&
        p.TranslatorEmail === TranslatorEmail &&
        p.TranslationLanguage === TranslationLanguage &&
        NUDGEABLE_STATUSES.has(p.DriftStatus as DriftStatus)
    );

    if (pagesToNudge.length === 0) {
      return;
    }

    setIsSending(true);
    setErrorMessage(undefined);

    try {
      await sendNudgeEmail(
        props.graphClient,
        TranslatorEmail,
        TranslatorName ?? '',
        pagesToNudge,
        SiteUrl
      );
      await markNudgeSent(
        props.sp,
        pagesToNudge.map((p) => p.PageGuid),
        TranslationLanguage
      );
      setIsSent(true);
      setShowSuccess(true);
      setTimeout(() => setShowSuccess(false), 3000);
      if (props.onNudgeSent) {
        props.onNudgeSent(props.item);
      }
    } catch {
      setErrorMessage('Failed to send nudge');
    } finally {
      setIsSending(false);
    }
  };

  if (isSending) {
    return <Spinner size={SpinnerSize.small} label="Sending..." labelPosition="right" />;
  }

  if (showSuccess) {
    return <span style={{ color: '#107c10', fontWeight: 600 }}>Sent!</span>;
  }

  const hasTranslator = !!props.item.TranslatorEmail;

  return (
    <DefaultButton
      text={isSent ? 'Sent' : 'Nudge'}
      disabled={isSent || !hasTranslator}
      title={!hasTranslator ? 'No translator configured for this page' : errorMessage}
      onClick={() => {
        void handleNudge();
      }}
    />
  );
}

export default NudgeButton;
