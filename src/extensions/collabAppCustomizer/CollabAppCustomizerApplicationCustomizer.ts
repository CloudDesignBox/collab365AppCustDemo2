import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import * as strings from 'CollabAppCustomizerApplicationCustomizerStrings';

const LOG_SOURCE: string = 'CollabAppCustomizerApplicationCustomizer';

//import styles as module
import styles from './mystyles.module.scss';

//not recommended
//require('./badstyles.scss');

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICollabAppCustomizerApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class CollabAppCustomizerApplicationCustomizer
  extends BaseApplicationCustomizer<ICollabAppCustomizerApplicationCustomizerProperties> {

    //placeholder
    private _headerPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // Call render method for generating the needed html elements
    this._renderPlaceHolders();

    return Promise.resolve<void>();
  }

  private _renderPlaceHolders(): void {
    //see which placeholders are available
    console.log('loading placeholders');
    console.log('Available placeholders: ',
    this.context.placeholderProvider.placeholderNames.map(name => PlaceholderName[name]).join(', '));

    // Handling the top placeholder
    if (!this._headerPlaceholder) {
      this._headerPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { //on dispose method
          });
          this._headerPlaceholder.domElement.innerHTML = `
          <div class="${styles.myheader}">
            <img src="https://clouddesignboxdev.sharepoint.com/sites/Team/SiteAssets/Collab365.png" alt="Collab365 Logo" />
              <div class="${styles.socialmediaholder}">
              <div class="${styles.socialtileholder}"><a class="${styles.socialtile} ${styles.facebook}" href="https://www.facebook.com"></a></div>
              <div class="${styles.socialtileholder}"><a class="${styles.socialtile} ${styles.twitter}" href="https://www.twitter.com"></a></div>
              <div class="${styles.socialtileholder}"><a class="${styles.socialtile} ${styles.linkedin}" href="https://www.linkedin.com"></a></div>
              <div class="${styles.socialtileholder}"><a class="${styles.socialtile} ${styles.youtube}" href="https://www.youtube.com"></a></div>
              <div class="${styles.socialtileholder}"><a class="${styles.socialtile} ${styles.instagram}" href="https://www.instagram.com"></a></div>
            </div>
          </div>
          `;
    }

  }
}
/**
 * use this url to test:
 * Replace "https://clouddesignboxdev.sharepoint.com/sites/team" with the URL of your test site
 * Team site
 * https://clouddesignboxdev.sharepoint.com/sites/team?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&customActions={"92a422b6-b906-42c9-9c40-685b7ddc6593":{"location":"ClientSideExtension.ApplicationCustomizer","properties":{"testMessage":"Hello as property!"}}}
 * Communication site
 * https://clouddesignboxdev.sharepoint.com/sites/communication?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&customActions={"92a422b6-b906-42c9-9c40-685b7ddc6593":{"location":"ClientSideExtension.ApplicationCustomizer","properties":{"testMessage":"Hello as property!"}}}
 */