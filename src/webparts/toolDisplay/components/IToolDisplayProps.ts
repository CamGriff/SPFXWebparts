import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IToolDisplayProps {
  context: WebPartContext;
  title: string;        // configurable via property pane
  columns: number;      // how many columns in the grid
}