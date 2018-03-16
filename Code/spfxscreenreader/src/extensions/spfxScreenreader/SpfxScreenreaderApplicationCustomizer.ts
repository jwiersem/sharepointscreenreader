import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';  

import * as strings from 'SpfxScreenreaderApplicationCustomizerStrings';

import styles from './AppCustomizer.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';

import { Screenreader } from './Screenreader/Screenreader';

const LOG_SOURCE: string = 'SpfxScreenreaderApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpfxScreenreaderApplicationCustomizerProperties {
  Top: string;
  apiUrl: string;
  autoPlay: boolean;
  selectors: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpfxScreenreaderApplicationCustomizer
  extends BaseApplicationCustomizer<ISpfxScreenreaderApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;

  private propertyNameApiUrl: string = "screenreader-apiUrlProperty";
  private propertynameAutoPlay: string = "screenreader-autoPlayProperty";

  private screenReader: Screenreader;

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // Added to handle possible changes on the existence of placeholders.
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    // Call render method for generating the HTML elements.
    this._renderPlaceHolders();
    return super.onInit().then(_ => {
      console.log("OnInit ran.");
    });
  }

  private async getProperties(aProperties: ISpfxScreenreaderApplicationCustomizerProperties): Promise<any>
  {
    return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/site/rootweb/lists/getByTitle('ScreenreaderSettings')/items?$select=screenreader_apiUrl,screenreader_autoPlay,screenreader_selectors&$top=1`,  
      SPHttpClient.configurations.v1)  
      .then((response: SPHttpClientResponse) => {  
        response.json().then((responseJSON: any) => {  
          console.log(responseJSON);
          
          if (responseJSON.value.length == 1)
          {
            var aItem = responseJSON.value[0];
            aProperties.apiUrl = aItem["screenreader_apiUrl"];
            aProperties.autoPlay = aItem["screenreader_autoPlay"];
            aProperties.selectors = aItem["screenreader_selectors"];

            console.log(aProperties.apiUrl);
            console.log(String(aProperties.autoPlay));
            console.log(aProperties.selectors);
            console.log('Properties set.');
            
          }
          else
          {
            console.log("Did not find single item");
            throw new Error("Did not find single item");
          }
        });  
      }); 
  }

  private async _renderPlaceHolders(): Promise<void> {

    await this.getProperties(this.properties).catch((err) => {
      console.log("Error getting properties from ScreenreaderSettings list. Does it exist?");
    });

    this.screenReader = new Screenreader(this.context.httpClient, this.properties);
    
    console.log('Available placeholders: ',
    this.context.placeholderProvider.placeholderNames.map(name => PlaceholderName[name]).join(', '));

    // Handling the top placeholder
    if (!this._topPlaceholder) {
      this._topPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { onDispose: this._onDispose });

      // The extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.log('The expected placeholder (Top) was not found.');
        return;
      }

      if (this.properties) {
        let topString: string = this.properties.Top;
        if (!topString) {
          topString = '(Top property was not defined.)';
        }

        if (this._topPlaceholder.domElement) {
          this._topPlaceholder.domElement.innerHTML = this.screenReader.render(styles);
        }     
           
        this.screenReader.addInteractivity();
      }
    }
  }

  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
