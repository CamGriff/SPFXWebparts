// Shared, in-memory selection store for the News Digest layout.
//
// Deliberately does NOT use PnP Modern Search's built-in itemSelectionProps
// / AsDataFilter mechanism — that feature is designed to connect two web
// parts together (select an item in one, filter another using one of its
// field values), not to collect a set of full item records for use outside
// the page. This store exists instead: a plain module-level Map, populated
// by a custom checkbox component on each card, read by the compose-digest
// dialog when the user is ready to send.
//
// Being module-level state (not tied to PnP's own render/selection state)
// means a selection made on page 1 of results survives paging to page 2,
// for as long as this page stays loaded — which is the more useful
// behaviour for "browse several pages, pick a few good ones" than PnP's
// own per-render selection would have given us.

export interface IDigestItem {
  id: string;           // normalized item identity (e.g. UniqueID, braces stripped)
  title: string;
  link: string;
  summary: string;
  thumbnailUrl: string | undefined;
}

type Listener = (items: IDigestItem[]) => void;

const selectedItems: Map<string, IDigestItem> = new Map();
const listeners: Set<Listener> = new Set();

function notify(): void {
  const snapshot = getSelectedItems();
  listeners.forEach((listener) => listener(snapshot));
}

/**
 * Adds an item to the digest selection, or updates it if already present
 * (e.g. re-selecting after a page's data reloaded with fresher fields).
 */
export function selectItem(item: IDigestItem): void {
  selectedItems.set(item.id, item);
  notify();
}

/**
 * Removes an item from the digest selection. Safe to call on an id that
 * isn't currently selected — no-ops rather than throwing.
 */
export function deselectItem(id: string): void {
  if (selectedItems.delete(id)) {
    notify();
  }
}

export function isSelected(id: string): boolean {
  return selectedItems.has(id);
}

export function getSelectedItems(): IDigestItem[] {
  return Array.from(selectedItems.values());
}

export function getSelectedCount(): number {
  return selectedItems.size;
}

export function clearSelection(): void {
  if (selectedItems.size === 0) {
    return;
  }
  selectedItems.clear();
  notify();
}

/**
 * Subscribes to selection changes (add/remove/clear). Used by the floating
 * "Compose digest" trigger to keep its live count in sync without needing
 * to re-render the whole card grid. Returns an unsubscribe function.
 */
export function onSelectionChange(listener: Listener): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}