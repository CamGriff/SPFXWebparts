export interface IUserPreference {
  id?: string;             // SP list item ID — needed for update vs create
  userId: string;          // UPN e.g. user@oecd.org
  tools: string[];         // ordered array of Tool IDs e.g. ["3","7","1"]
  lastUpdated: string;
}