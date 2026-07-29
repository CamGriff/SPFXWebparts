import * as React from 'react';
import { useTheme, Text, mergeStyles } from '@fluentui/react';
import { DriftStatus } from '../../../../models/DriftStatus';
import { IDriftItem } from '../../../../models/IDriftItem';

export type SummaryTileKey = DriftStatus | 'All';

export interface ISummaryTilesProps {
  items: IDriftItem[];
  activeFilter: SummaryTileKey;
  onFilterChange: (filter: SummaryTileKey) => void;
}

interface ITileDefinition {
  key: SummaryTileKey;
  label: string;
}

const TILE_DEFINITIONS: ITileDefinition[] = [
  { key: 'All', label: 'Total' },
  { key: DriftStatus.InSync, label: 'In Sync' },
  { key: DriftStatus.Stale, label: 'Stale' },
  { key: DriftStatus.Missing, label: 'Missing' },
  { key: DriftStatus.Orphaned, label: 'Orphaned' },
  { key: DriftStatus.Abandoned, label: 'Abandoned' }
];

function countForTile(items: IDriftItem[], key: SummaryTileKey): number {
  if (key === 'All') {
    return items.length;
  }
  return items.filter((item) => item.DriftStatus === key).length;
}

function SummaryTiles(props: ISummaryTilesProps): JSX.Element {
  const theme = useTheme();

  const containerClass = mergeStyles({
    display: 'flex',
    gap: '12px',
    flexWrap: 'wrap',
    marginBottom: '16px'
  });

  return (
    <div className={containerClass}>
      {TILE_DEFINITIONS.map((tile) => {
        const isActive = props.activeFilter === tile.key;
        const count = countForTile(props.items, tile.key);
        const countColor =
          isActive && tile.key !== 'All' ? theme.palette.themePrimary : theme.semanticColors.bodyText;

        const tileClass = mergeStyles({
          cursor: 'pointer',
          padding: '12px 20px',
          borderRadius: theme.effects.roundedCorner4,
          minWidth: '100px',
          textAlign: 'center',
          border: isActive
            ? `2px solid ${theme.palette.themePrimary}`
            : `1px solid ${theme.palette.neutralLight}`,
          backgroundColor: isActive ? theme.palette.themeLighterAlt : theme.semanticColors.bodyBackground,
          boxShadow: isActive ? 'none' : theme.effects.elevation4,
          outline: 'none',
          fontFamily: 'inherit',
          selectors: {
            ':hover': {
              borderColor: theme.palette.themePrimary,
              backgroundColor: theme.palette.themeLighterAlt
            },
            ':focus-visible': {
              outline: `2px solid ${theme.palette.themePrimary}`,
              outlineOffset: '2px'
            }
          }
        });

        return (
          <button
            key={tile.key}
            type="button"
            className={tileClass}
            onClick={() => props.onFilterChange(tile.key)}
          >
            <Text
              variant="xxLarge"
              styles={{ root: { display: 'block', fontWeight: 600, color: countColor } }}
            >
              {count}
            </Text>
            <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary } }}>
              {tile.label}
            </Text>
          </button>
        );
      })}
    </div>
  );
}

export default SummaryTiles;
