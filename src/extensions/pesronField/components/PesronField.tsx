import { Log } from "@microsoft/sp-core-library";
import * as React from "react";
import {
  SPHttpClient,
  SPHttpClientConfiguration,
  SPHttpClientResponse,
  ODataVersion,
  ISPHttpClientConfiguration,
} from "@microsoft/sp-http";
import { ServiceScope } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { LivePersona } from "@pnp/spfx-controls-react/lib/LivePersona";
import styles from "./PesronField.module.scss";
import { LivePersonaCard } from '../../../controls/LivePersonaCard';
import { Persona } from "office-ui-fabric-react";
export interface IPesronFieldProps {

  text: string;
  imageUrl: string;
  imageInitials: string;
  Title: string;
  jobTitle: string;
  email: string;
  optionalText: string;
  objet: any;
  serviceScope: any
  context: WebPartContext
}

const LOG_SOURCE: string = "PesronField";

export default class PesronField extends React.Component<
  IPesronFieldProps,
  {}
> {
  context: WebPartContext
  constructor(props: IPesronFieldProps) {
    super(props);

    //this.getpro(this.props.email);
  }

  public componentDidMount(): void {
    Log.info(LOG_SOURCE, "React Element: PesronField mounted");
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, "React Element: PesronField unmounted");
  }

  /*<div className={styles.lpcSample}>

       <LivePersonaCard
        user={{
          displayName: PreferredName,
          email: email,
        }}
        serviceScope={this.props.serviceScope}
       

/>
       

        {PreferredName}
        {telephone.length > 0 && <p className={styles.p}>Tell: {telephone}</p>}

        {Department.length > 0 && <p className={styles.p}>Department: {Department}</p>}
        {option.length > 0 && <p className={styles.p}>option: {option}</p>}
        {email.length > 0 && <p className={styles.p}> {email}</p>}


      </div>*/
  public render(): React.ReactElement<{}> {

    //console.log(this.props.serviceScope)
    const telephone = this.props.objet.cellPhone;
    const PreferredName = this.props.objet.PreferredName;
    const Department = this.props.objet.Department;
    const option = this.props.objet.option;
    const email = this.props.objet.email;
    const PictureURL = this.props.objet.PictureURL
    //console.log(PictureURL)

    return (
      <div id="thisid" className={styles.cell}>
        <LivePersona upn={email}
          template={
            <>
              <div className={styles.chip}>
                <img src={PictureURL}></img>
                {PreferredName}<br></br>{telephone}<br></br>{Department}</div>
            </>
          }
          serviceScope={this.props.serviceScope}
        />

      </div>
    );
  }
}
