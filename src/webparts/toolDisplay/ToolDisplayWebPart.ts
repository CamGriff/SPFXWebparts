import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IToolDisplayProps } from "./components/IToolDisplayProps";
import ToolDisplay from "./components/ToolDisplay";

export interface IToolDisplayWebPartProps {
  title: string;
  columns: number;
}

export default class ToolDisplayWebPart extends BaseClientSideWebPart<IToolDisplayWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IToolDisplayProps> = React.createElement(
      ToolDisplay,
      {
        context: this.context,
        title: this.properties.title || "My Tools",
        columns: this.properties.columns || 4,
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Tool Display Settings" },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField("title", {
                  label: "Webpart Title",
                }),
                PropertyPaneSlider("columns", {
                  label: "Number of Columns",
                  min: 2,
                  max: 6,
                  step: 1,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}