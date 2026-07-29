export interface ITool {
  id: string;
  titleEN: string;
  titleFR: string;
  descriptionEN: string;
  descriptionFR: string;
  urlEN: string;
  urlFR: string;
  toolIcon: string;
  // Resolved at runtime based on user language
  title?: string;
  description?: string;
  url?: string;
}