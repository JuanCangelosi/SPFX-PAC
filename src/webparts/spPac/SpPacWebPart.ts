import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";

import styles from "./SpPacWebPart.module.scss";
import * as strings from "SpPacWebPartStrings";

require("./pac-element.js");
require("./pac-styles.css");

export interface ISpPacWebPartProps {
  description: string;
  name: string;
  userList: string;
  managerList: string;
  finalList: string;
}

export default class SpPacWebPart extends BaseClientSideWebPart<ISpPacWebPartProps> {

  public render(): void {
    console.log(this.properties);
    this.domElement.innerHTML = `
      <div class="${ styles.spPac}">
        <div class="${ styles.container}">
        <pac-element [configurations]="${this.properties}"></pac-element>
        <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "PAC app configurations"
          },
          groups: [
            {
              groupName: "Configurations",
              groupFields: [
                PropertyPaneTextField("name", {
                  label: "name",
                  value: "PAC",
                }),
                PropertyPaneTextField("userList", {
                  label: "User Requests List",
                  value: "PACRequest"
                }),
                PropertyPaneTextField("managerList", {
                  label: "Responses List",
                  value: "PACResponse"
                }),
                PropertyPaneTextField("finalList", {
                  label: "Final Request List",
                  value: "PACRequestFinal"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
