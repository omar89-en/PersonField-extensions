import * as React from "react";
import * as ReactDOM from "react-dom";

import { Log } from "@microsoft/sp-core-library";
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters,
} from "@microsoft/sp-listview-extensibility";

import * as strings from "PesronFieldFieldCustomizerStrings";
import PesronField, { IPesronFieldProps } from "./components/PesronField";
import {
  SPHttpClient,
  SPHttpClientConfiguration,
  SPHttpClientResponse,
  ODataVersion,
  ISPHttpClientConfiguration,
} from "@microsoft/sp-http";

/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPesronFieldFieldCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
  cellphinee: string;
}

const LOG_SOURCE: string = "PesronFieldFieldCustomizer";

export default class PesronFieldFieldCustomizer extends BaseFieldCustomizer<IPesronFieldFieldCustomizerProperties> {
  public onInit(): Promise<void> {
    // Add your custom initialization to this method.  The framework will wait
    // for the returned promise to resolve before firing any BaseFieldCustomizer events.
    Log.info(
      LOG_SOURCE,
      "Activated PesronFieldFieldCustomizer with properties:"
    );
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    Log.info(
      LOG_SOURCE,
      `The following string should be equal: "PesronFieldFieldCustomizer" and "${strings.Title}"`
    );
    return Promise.resolve();
  }

  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
 
 this.getProperties(event.fieldValue[0].email, event);

   
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }

  public getProperties(url: any, event: IFieldCustomizerCellEventParameters) {
    
    var libraries: any[] = [];
    let cellPhone: string;
    let Department: string;
    let Title: string;
    let option: string;
    let PreferredName: string;
    let email:string
let PictureURL:string;
    var initials = event.fieldValue[0].title.match(/\b\w/g) || [];
    initials = (
      (initials.shift() || "") + (initials.pop() || "")
    ).toUpperCase();

    this.context.spHttpClient
      .get(
        `${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='i:0%23.f|membership|${url}'`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        response.json().then((PropertyValues: any) => {
         
          //console.log(PropertyValues.UserProfileProperties);
          PropertyValues.UserProfileProperties.forEach((item) => {
            if (item.Key == "WorkPhone") {
              cellPhone = item.Value;
            }
            if (item.Key == "Department") {
              Department = item.Value;
            }
            if (item.Key == "SPS-EmailOptin") {
              option = item.Value;
            }
            
            if (item.Key == "SPS-JobTitle") {
              Title = item.Value;
            }

            if (item.Key == "PreferredName") {
              PreferredName = item.Value;
            }
            if (item.Key == "UserName") {
              email = item.Value;
            }
            if (item.Key == "PictureURL") {
              PictureURL = item.Value;
            }
            
          });
          

          const pesronField: React.ReactElement<{}> = React.createElement(
            PesronField,
            {
              serviceScope: this.context.serviceScope,
              imageUrl: event.fieldValue[0].picture,
              imageInitials: initials,
              Title: event.fieldValue[0].title,
              jobTitle: event.fieldValue[0].jobTitle,
              email: event.fieldValue[0].email,
              optionalText: "",
              objet: {
                cellPhone: cellPhone,
                Department: Department,
                Title: Title,
                option:option,
                PreferredName: PreferredName,
                email:email,
                PictureURL: PictureURL,
              },
            } as IPesronFieldProps
          );

          ReactDOM.render(pesronField, event.domElement);
      
        });
      });
   
  }
}
