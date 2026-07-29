import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IToolPickerProps } from "./components/IToolPickerProps";
import ToolPicker from "./components/ToolPicker";

export default class ToolPickerWebPart extends BaseClientSideWebPart<{}> {

  public render(): void {
    const element: React.ReactElement<IToolPickerProps> = React.createElement(
      ToolPicker,
      {
        context: this.context,
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
      pages: []
    };
  }
}