import * as React from 'react';
import { useTheme, mergeStyles } from '@fluentui/react';
import { DriftStatus } from '../../../../models/DriftStatus';

export interface IStatusPillProps {
  status: DriftStatus;
}

function StatusPill(props: IStatusPillProps): JSX.Element {
  const theme = useTheme();

  const colorMap: Record<DriftStatus, { background: string; color: string }> = {
    [DriftStatus.InSync]: {
      background: '#dff6dd',
      color: theme.palette.greenDark
    },
    [DriftStatus.Stale]: {
      background: theme.semanticColors.warningBackground,
      color: theme.semanticColors.warningText
    },
    [DriftStatus.Missing]: {
      background: theme.semanticColors.errorBackground,
      color: theme.semanticColors.errorText
    },
    [DriftStatus.Orphaned]: {
      background: '#fde0cc',
      color: theme.palette.orange
    },
    [DriftStatus.Abandoned]: {
      background: theme.palette.neutralLighter,
      color: theme.palette.neutralSecondary
    }
  };

  const colors = colorMap[props.status];

  const pillClass = mergeStyles({
    display: 'inline-block',
    padding: '2px 10px',
    borderRadius: '10px',
    fontSize: '12px',
    fontWeight: 600,
    whiteSpace: 'nowrap',
    backgroundColor: colors.background,
    color: colors.color
  });

  return <span className={pillClass}>{props.status}</span>;
}

export default StatusPill;
