import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import * as strings from 'AnalyticsApplicationCustomizerStrings';
import { ISiteUsers, sp } from '@pnp/sp/presets/all';
import { Title } from 'AnalyticsApplicationCustomizerStrings';
import * as $ from 'jquery';
const LOG_SOURCE: string = 'AnalyticsApplicationCustomizer';


/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAnalyticsApplicationCustomizerProperties {
  // This is an example; replace with your own property
  trackingID: string;
  context?: WebPartContext;

}
export interface IAnalyticsApplicationCustomizerState {
  User: any;

}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AnalyticsApplicationCustomizer extends BaseApplicationCustomizer<IAnalyticsApplicationCustomizerProperties> {


  private currentPage = "";

  /** Statement for trigger ony once the initialization of analytics script
   * @private
   */
  private isInitialLoad = true;

  /** Get Current page URL
   * @returns URL of the current Page
   * @private
   */
  private getFreshCurrentPage(): string {
    return window.location.pathname + window.location.search;
  }

  /** Update current page navigation
   * @private
   */
  private updateCurrentPage(): void {
    this.currentPage = this.getFreshCurrentPage();
  }

  /** Navigation and search event
   * @private
   */
  private navigatedEvent(): void {
    debugger;
    console.log("navigatedEvent");
    let trackingID: string = this.properties.trackingID;
    if (!trackingID) {
      Log.info(LOG_SOURCE, `${strings.MissingID}`);
    } else {
      console.log("trackingID false");

      Log.info(LOG_SOURCE, `Tracking Site ID: ${trackingID}`);
      const navigatedPage = this.getFreshCurrentPage();

      if (this.isInitialLoad) {
        console.log("this.isInitialLoad true");
        Log.info(LOG_SOURCE, `Initial load`);
        this.realInitialNavigatedEvent(trackingID);
        // this.realNavigatedEvent(trackingID);
        this.updateCurrentPage();

        this.isInitialLoad = false;

      } else if (!this.isInitialLoad && navigatedPage !== this.currentPage) {
        console.log("this.isInitialLoad false");
        Log.info(LOG_SOURCE, `Not initial load`);
        // this.realNavigatedEvent(trackingID);
        this.updateCurrentPage();
      }
    }
  }

  /** Inital Page load - init analytics
   * @param trackingID Google Analytics Tracking Site ID
   * @private
   */
  private realInitialNavigatedEvent(trackingID: string): void {
    let DepartmentName: string = this._getUserProfileProperties();
    console.log("DepartmentNameResult", DepartmentName);
    debugger;
    console.log("realInitialNavigatedEvent");
    Log.info(LOG_SOURCE, `Tracking full page load...`);
    const dimensionValue = this.context.pageContext.legacyPageContext['userLoginName'];
    console.log(" this.context.pageContext.legacyPageContext", this.context.pageContext.legacyPageContext);
    const dimensionValueUserName = this.context.pageContext.legacyPageContext['userDisplayName'];
    console.log(" dimensionValueUserName", dimensionValueUserName);
    var gtagScript = document.createElement("script");
    gtagScript.type = "text/javascript";
    gtagScript.src = `https://www.googletagmanager.com/gtag/js?id=${trackingID}`;
    gtagScript.async = true;
    document.head.appendChild(gtagScript);

    eval(`
             window.dataLayer = window.dataLayer || [];
             function gtag(){dataLayer.push(arguments);}
             gtag('js', new Date());
             gtag('config', 'G-5ZTYPWR9W8');
             gtag('config',  '${trackingID}',{'custom_map':{'dimension1':'UserID'}});
             gtag('event','pageview',{'UserID':'${dimensionValue}'});
             gtag('config',  '${trackingID}',{'custom_map':{'dimension2':'Username'}});
             gtag('event','pageview',{'Username':'${dimensionValueUserName}'});
             gtag('config',  '${trackingID}',{'custom_map':{'dimension3':'Department'}});
             gtag('event','pageview',{'Department':'${DepartmentName}'});
           `);
  }

  public _getUserProfileProperties(): string {
    console.log("_getUserProfileProperties");
    var DepartmentName: string = "";
    var reactHandler = this;
    var url = "https://sonorasoftware0.sharepoint.com" + "/_api/SP.UserProfiles.PeopleManager/GetMyProperties";
    $.ajax({
      url: url,
      method: "GET",
      headers: {
        "Accept": "application/json; odata=verbose"

      }, async: false,
      success: function (data) {
        console.log("User properties", data);
        var Department: string = data.d.UserProfileProperties.results.map((dt) => {
          if (dt.Key == "SPS-Department") DepartmentName = dt.Value;
        });
        console.log("Department", Department);
        const deptname: any = Department;
        console.log("DepartmentName", DepartmentName);
      },
      error: function (err) { console.log(err); }
    });
    return DepartmentName;
  }


  /** Partial Page load
   * @param trackingID Google Analytics Tracking Site ID
   * @private 
   */
  private realNavigatedEvent(trackingID: string): void {

    Log.info(LOG_SOURCE, `Tracking partial page load...`);
    eval(`
   
          if(ga) {
             ga('create', '${trackingID}', 'auto');
            ga('send', 'pageview', { 'dimension1': this.context.pageContext.legacyPageContext['userLoginName'] });
           
          }
`);

    console.log(" this.context.pageContext.legacyPageContext['userDisplayName']", this.context.pageContext.legacyPageContext['userDisplayName']);

  }


  @override
  public onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context
    });
    Log.info(LOG_SOURCE, `Initialized Google Analytics`);
    /* This event is triggered when user performed a search from the header of SharePoint */
    this.context.placeholderProvider.changedEvent.add(
      this,
      this.navigatedEvent
    );
    /* This event is triggered when user navigate between the pages */
    this.context.application.navigatedEvent.add(this, this.navigatedEvent);

    return Promise.resolve();
  }
}
