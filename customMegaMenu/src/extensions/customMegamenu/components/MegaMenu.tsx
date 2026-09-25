import * as React from 'react';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import styles from './MegaMenu.module.scss';

export interface IMegaMenuProps {
  context: ApplicationCustomizerContext;
}

// The shape the component renders from. Built in-memory from the flat list
// rows — the component below never sees the list's raw structure.
interface IMenuNode {
  id: number;
  label: string;
  url: string;
  openInNewTab: boolean;
  children: IMenuNode[];
}

// Raw row shape as it comes back from the list. Note ParentId is a lookup,
// so it arrives as an object ({ Id, Title }) or undefined, NOT a number —
// this is the single most common thing to get wrong reading a self-lookup.
interface IRawMenuItem {
  Id: number;
  Title: string;
  NavUrl?: { Url: string } | string;
  Levels: string;            // choice column, internal name "Levels" (Level was taken)
  SortOrder?: number;
  OpenInNewTab?: boolean;
  ParentId?: { Id: number };
}

const LIST_TITLE = 'MegaMenuItems';

export const MegaMenu: React.FunctionComponent<IMegaMenuProps> = (props) => {
  const [tree, setTree] = React.useState<IMenuNode[]>([]);
  const [loaded, setLoaded] = React.useState<boolean>(false);

  React.useEffect(() => {
    const load = async (): Promise<void> => {
      try {
        const sp = spfi().using(SPFx(props.context));

        const items: IRawMenuItem[] = await sp.web.lists
          .getByTitle(LIST_TITLE)
          .items
          .select('Id', 'Title', 'NavUrl', 'Levels', 'SortOrder', 'OpenInNewTab', 'ParentId/Id')
          .expand('ParentId')
          .orderBy('SortOrder', true)
          .top(500)();

        setTree(buildTree(items));
      } catch (error) {
        // Fail quietly to an empty menu rather than throwing on every page
        // load — a broken menu shouldn't take the whole page down with it.
        console.error('[MegaMenu] Failed to load menu items:', error);
        setTree([]);
      } finally {
        setLoaded(true);
      }
    };

    void load();
  }, [props.context]);

  if (!loaded) {
    // Render nothing until loaded — avoids a flash of empty bar. Kept
    // deliberately minimal; a skeleton could go here later if wanted.
    return null;
  }

  return (
    <div className={styles.megaMenu}>
      <nav className={styles.nav}>
        {tree.map((item) => (
          <div key={item.id} className={styles.navItem}>
            <a
              href={item.url || '#'}
              className={styles.navLink}
              target={item.openInNewTab ? '_blank' : undefined}
              rel={item.openInNewTab ? 'noopener noreferrer' : undefined}
            >
              {item.label}
              {item.children.length > 0 && (
                <span className={styles.chevron}> &#9662;</span>
              )}
            </a>

            {item.children.length > 0 && (
              <div className={styles.flyout}>
                {item.children.map((child) => (
                  <div key={child.id} className={styles.flyoutColumn}>
                    <a
                      href={child.url || '#'}
                      className={styles.flyoutLink}
                      target={child.openInNewTab ? '_blank' : undefined}
                      rel={child.openInNewTab ? 'noopener noreferrer' : undefined}
                    >
                      {child.label}
                    </a>

                    {child.children.length > 0 && (
                      <div className={styles.flyoutSubList}>
                        {child.children.map((grandchild) => (
                          <a
                            key={grandchild.id}
                            href={grandchild.url || '#'}
                            className={styles.flyoutSubLink}
                            target={grandchild.openInNewTab ? '_blank' : undefined}
                            rel={grandchild.openInNewTab ? 'noopener noreferrer' : undefined}
                          >
                            {grandchild.label}
                          </a>
                        ))}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            )}
          </div>
        ))}
      </nav>
    </div>
  );
};

// Turns the flat list rows into a three-level tree, keyed on the
// ParentId lookup's .Id. Items are already sorted by SortOrder from the
// query, so sibling order is preserved by insertion.
function buildTree(items: IRawMenuItem[]): IMenuNode[] {
  const nodesById: Map<number, IMenuNode> = new Map();

  // First pass: create a node for every row.
  for (const item of items) {
    nodesById.set(item.Id, {
      id: item.Id,
      label: item.Title,
      url: normalizeUrl(item.NavUrl),
      openInNewTab: !!item.OpenInNewTab,
      children: []
    });
  }

  const roots: IMenuNode[] = [];

  // Second pass: attach each node to its parent, or treat it as a root.
  for (const item of items) {
    const node = nodesById.get(item.Id);
    if (!node) {
      continue;
    }

    const parentId = item.ParentId?.Id;
    if (parentId && nodesById.has(parentId)) {
      nodesById.get(parentId)!.children.push(node);
    } else {
      // No parent (or a parent that isn't in the set) => top-level item.
      roots.push(node);
    }
  }

  return roots;
}

// A URL column comes back as { Url: '...' }; guard for the string case and
// for it being absent entirely (a flyout-only parent has no URL).
function normalizeUrl(value: { Url: string } | string | undefined): string {
  if (!value) {
    return '';
  }
  if (typeof value === 'string') {
    return value;
  }
  return value.Url || '';
}