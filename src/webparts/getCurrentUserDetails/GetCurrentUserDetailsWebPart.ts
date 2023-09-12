import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';
//import * as pnp from "sp-pnp-js";
import styles from './GetCurrentUserDetailsWebPart.module.scss';
import * as strings from 'GetCurrentUserDetailsWebPartStrings';
import * as $ from 'jquery';
import {ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
//import { sp } from "@pnp/sp";  
//import "@pnp/sp/profiles"; 
//declare var __userEmail__ :;
export interface IGetCurrentUserDetailsWebPartProps {
  description: string;
}

export default class GetCurrentUserDetailsWebPart extends BaseClientSideWebPart<IGetCurrentUserDetailsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
   
  public _getUserLocale()  {
    
      var siteUrl = this.context.pageContext.web.absoluteUrl;
      console.log(siteUrl);
      $.ajax({
          url: siteUrl + "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
          method: "GET",
          headers: { "Accept": "application/json; odata=verbose" },
          success: function (data) {
           // var items = data.d.UserProfileProperties.results[112]
            
              console.log(data);
              document.getElementById('spUserProfilePropertiesEmployeeID')!.innerHTML = data.d.UserProfileProperties.results[112].Value;
              document.getElementById('spUserProfilePropertiesEmail')!.innerHTML = data.d.Email;
              //const userEmail = data.d.Email
              globalThis.empID = data.d.UserProfileProperties.results[112].Value;
              globalThis.userEmail = data.d.Email;

          },
          error: function(error) {
              console.log(error);
          }
      });
      
      }


  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.getCurrentUserDetails} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
       <!-- <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" /> -->
        <h2>Well done!, ${escape(this.context.pageContext.user.displayName)}!</h2>
       <!-- <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div> -->
      </div>
      <div>
      <p><label for="email">Email:</label><span id = "spUserProfilePropertiesEmail"/></span</p>
      <p><label for="empID">Employee ID:</label><span id="spUserProfilePropertiesEmployeeID" /> </span> </p>
      <input type="button" id="BttnEmp" value="Save">

      </div>
    </section>`;

    this._getUserLocale();
    this._bindsave();

    //this.loadUserDetails();
    //this.GetUserDetails();
  }

private _bindsave(): void{
this.domElement.querySelector('#BttnEmp')?.addEventListener('click',() => {this.addListItem();
});
}

private addListItem() : void {


  //var empID = document.getElementById("spUserProfilePropertiesEmployeeID");
  //var empEmail = document.getElementById('spUserProfilePropertiesEmail');
  const siteurl: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('GetCurrentUserDetails')/items"
  const itemBody : any = {
    "Title" : globalThis.userEmail,
    "EmployeeID" : globalThis.empID
   

  };
  console.log('EmployeeID');
  console.log('Title');
  const spHttpClientOptions : ISPHttpClientOptions = {

    "body" : JSON.stringify(itemBody)
  };

  this.context.spHttpClient.post(siteurl, SPHttpClient.configurations.v1, spHttpClientOptions).then((response: SPHttpClientResponse) =>
  {

    alert('Success');
  });
}
 

  protected  onInit(): Promise<void> {

    //await super.onInit();  
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
    //sp.setup(this.context); 
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
