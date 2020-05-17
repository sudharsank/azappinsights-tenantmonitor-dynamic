import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'AzureApplicationInsightsTenantMonitorApplicationCustomizerStrings';

import {
  ApplicationInsights,
  ITelemetryItem,
  SeverityLevel,
  Snippet
} from "@microsoft/applicationinsights-web";
import jsSHA from "jssha";

const LOG_SOURCE: string = 'AzureApplicationInsightsTenantMonitorApplicationCustomizer';

export interface IAzureApplicationInsightsTenantMonitorApplicationCustomizerProperties {
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AzureApplicationInsightsTenantMonitorApplicationCustomizer
  extends BaseApplicationCustomizer<IAzureApplicationInsightsTenantMonitorApplicationCustomizerProperties> {

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // load App Insights
    const appInsights = await this._initAppInsights();

    // track page view
    this._trackPageView(appInsights);

    // attach to buttons
    this._attachLogButtonClick(appInsights, 'O365_MainLink_NavMenu', 'SPO_UX|SUITE_BAR|WAFFLE_MENU_CLICK', 'Waffle Menu');
    this._attachLogButtonClick(appInsights, 'TipsNTricksButton', 'SPO_UX|SUITE_BAR|TIPSTRICKS_BUTTON_CLICK', 'Tips & tricks icon');
    this._attachLogButtonClick(appInsights, 'O365_MainLink_Bell', 'SPO_UX|SUITE_BAR|NOTIFICATION_BUTTON_CLICK', 'Notifications icon');
    this._attachLogButtonClick(appInsights, 'O365_MainLink_Settings', 'SPO_UX|SUITE_BAR|SETTINGS_BUTTON_CLICK', 'Settings icon');
    this._attachLogButtonClick(appInsights, 'O365_MainLink_Help', 'SPO_UX|SUITE_BAR|HELP_BUTTON_CLICK', 'Help icon');
    this._attachLogButtonClick(appInsights, 'O365_MainLink_Me', 'SPO_UX|SUITE_BAR|PROFILE_BUTTON_CLICK', 'Profile icon');

    // fire randomized exception
    this._raiseRandomizedException(appInsights);

    return Promise.resolve();
  }

  private _initAppInsights(): Promise<ApplicationInsights> {
    // setup AI configuration
    const aiconfig = {
      instrumentationKey: AZURE_APPINSIGHTS_INSTRUMENTATION_KEY,
      disableAjaxTracking: true,
      disableFetchTracking: false,
      enableCorsCorrelation: true,
      maxBatchInterval: 0,        // good to set when debugging, but consider omitting in prod for perf
      namePrefix: 'spo-tenant-global'
    };

    // init & load app insights
    const appInsightsInstance = new ApplicationInsights(<Snippet>{ config: aiconfig });
    appInsightsInstance.loadAppInsights();

    // set the version of this script
    appInsightsInstance.context.application.ver = AZURE_APPINSIGHTS_APP_VERSION;

    // set auth context = logged in user
    //  REC: don't store PII in Azure App Insights
    //  REC: 1-way hash the user login & store as context
    const SHA256 = new jsSHA('SHA-256', 'TEXT');
    SHA256.update(this.context.pageContext.user.loginName);
    appInsightsInstance.setAuthenticatedUserContext(SHA256.getHash('HEX'));

    // add name & solution details of the app to every request
    appInsightsInstance.addTelemetryInitializer((envelope: ITelemetryItem) => {
      envelope.data['app.name'] = AZURE_APPINSIGHTS_APP_NAME;
      envelope.data['app.solution_id'] = AZURE_APPINSIGHTS_SOLUTION_ID;
      envelope.data['app.solution_version'] = AZURE_APPINSIGHTS_SOLUTION_VERSION;
    });

    return Promise.resolve(appInsightsInstance);
  }

  /*
   * Log current page view.
  */
  private _trackPageView(appInsights: ApplicationInsights): void {
    appInsights.trackPageView({
      properties: {
        site_id: this.context.pageContext.site.id,
        web_id: this.context.pageContext.web.id,
        web_title: this.context.pageContext.web.title
      }
    });
  }

  /*
   * Log AppInsights event to button clicks.
  */
  private _attachLogButtonClick(appInsights: ApplicationInsights, element_id: string, event_id: string, element_name: string): void {
    if (!document.getElementById(element_id)) {
      // most common reason: element not on page
      appInsights.trackTrace({
        message: `Pausing 10s to attach to '${element_name}'...`,
        severityLevel: SeverityLevel.Verbose
      });
      setTimeout(() => {
        // second attempt
        try {
          document.getElementById(element_id)
            .addEventListener('click', () => {
              appInsights.trackEvent({ name: event_id });
            });
          appInsights.trackTrace({
            message: `Attached to '${element_name}' 'click' event...`,
            severityLevel: SeverityLevel.Verbose
          });
        } catch (error) {
          appInsights.trackTrace({
            message: `Error attaching 'click' event to ${element_name} on 2nd attempt...`,
            severityLevel: SeverityLevel.Warning,
            properties: {
              target_element_id: element_id,
              error_name: error.name,
              error_message: error.message,
              error: JSON.stringify(error)
            }
          });
        }
      }, 10000);
    }
  }

  /*
   * Raise a randomized exception.
  */
  private _raiseRandomizedException(appInsights: ApplicationInsights): void {
    // if the current minute is evenly divisible by 7...
    if ((new Date).getMinutes() % 7 === 0) {
      // force an exception to happen
      try {
        // access something that doesn't exist to force error....
        let answer = (document as any).answerTo.theUniverse;
      } catch (error) {
        // log exception
        appInsights.trackException({
          exception: error,
          properties: {
            message: "Error accessing a property (theUniverse) on an undefined property (document.answerTo)"
          },
          severityLevel: SeverityLevel.Error
        });
      }
    }
  }
}
