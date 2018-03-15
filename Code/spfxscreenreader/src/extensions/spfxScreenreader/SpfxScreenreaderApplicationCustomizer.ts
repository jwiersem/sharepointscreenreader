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

import { sp } from "@pnp/sp";

import { ScreenreaderService,IScreenreaderServiceConfiguration } from './Services/ScreenreaderService';

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
  private atextToSpeechService: ScreenreaderService;
  private allText: string[] = [];

  private propertyNameApiUrl: string = "screenreader-apiUrlProperty";
  private propertynameAutoPlay: string = "screenreader-autoPlayProperty";

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // Added to handle possible changes on the existence of placeholders.
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    // Call render method for generating the HTML elements.
    this._renderPlaceHolders();
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  private scrapePage(): string[]
  {
    let allTextToRead: string[] = [];

    let heroElements: HTMLCollectionOf<Element> =  document.getElementsByClassName('ms-FocusZone');

    for(var i = 0; i < heroElements.length; i++)
    {
      var aText: string = heroElements[i].getAttribute('aria-label');

      if (aText)
      {
        console.log(aText);
        allTextToRead.push(aText);
      }
      else{
        console.log('No aria-label found for this element.');
      }
    }

    return allTextToRead;
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
            console.log(aProperties.autoPlay);
            console.log(aProperties.selectors);
            console.log('Properties set.');

          }
          else{
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

    if (!this.properties.apiUrl || 0 === this.properties.apiUrl.length)
    {
      this.properties.apiUrl = "https://prod-32.westeurope.logic.azure.com:443/workflows/737b64d81a9e4dc8b0dd1b938789df2b/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=PQbLDWnyX2m7GK6wRF5p03dvPacpYYPh87h5jp352dM";
    }

    let config: IScreenreaderServiceConfiguration = {
      httpClient: this.context.httpClient,
      apiUrl: this.properties.apiUrl
    };

    this.atextToSpeechService = new ScreenreaderService(config);

    console.log('HelloWorldApplicationCustomizer._renderPlaceHolders()');
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
        console.error('The expected placeholder (Top) was not found.');
        return;
      }

      if (this.properties) {
        let topString: string = this.properties.Top;
        if (!topString) {
          topString = '(Top property was not defined.)';
        }

        if (this._topPlaceholder.domElement) {
          this._topPlaceholder.domElement.innerHTML = `
        <div class="${styles.app}">
          <div id="screenreader-settings-wrapper" class="screenreader-settings-wrapper" style="display:none;">
            <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.top}">          
              <input id="input-apiUrl" type="text" value="${this.properties.apiUrl}"></input>
              <input id="input-autoPlay" type="checkbox" id="autoPlay" name="autoPlay" value="${this.properties.autoPlay}"></input>
              <label for="autoPlay">Autoplay?</label>
              <i id="screenreader-save-operation" class="screenreader-settings ms-Icon ms-Icon--Save x-hidden-focus" title="Save" aria-hidden="true" data-reactid=".0.0.$=10.0.1.$=1$3.6.1.$2.0"></i>
            </div>
          </div>
          <div id="screenreader-audioplayer-wrapper" class="screenreader-audioplayer-wrapper">
            <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.top}">

              <i class="ms-Icon ms-Icon--Rewind x-hidden-focus" title="Rewind" aria-hidden="true" data-reactid=".0.0.$=10.0.1.$=1$3.6.1.$0.0"></i>

              <i class="ms-Icon ms-Icon--CircleStopSolid x-hidden-focus" title="CircleStopSolid" aria-hidden="true" data-reactid=".0.0.$=10.0.1.$=1$3.6.1.$1.0"></i>

              <i class="ms-Icon ms-Icon--Play x-hidden-focus" title="Play" aria-hidden="true" data-reactid=".0.0.$=10.0.1.$=1$3.6.1.$1.0"></i> 
                        
              Reading screen!

              <i class="screenreader-settings ms-Icon ms-Icon--EditSolidMirrored12 x-hidden-focus" title="EditSolidMirrored12" aria-hidden="true" data-reactid=".0.0.$=10.0.1.$=1$3.6.1.$11.0"></i>
            
            </div>
          </div>          
        </div>
        `;
        }     
           
        let self = this;

        var settingsElements = document.getElementsByClassName('screenreader-settings');
      
        for(var i = 0;i < settingsElements.length; i++)
        {
          settingsElements[i].addEventListener('click', function()
          {
            self.toggle(document.getElementById('screenreader-settings-wrapper'));
            self.toggle(document.getElementById('screenreader-audioplayer-wrapper'));
          });
        }

        var saveElement = document.getElementById('screenreader-save-operation');

        saveElement.addEventListener('click', function(){
          var apiUrlElement = document.getElementById('input-apiUrl');
          var autoPlayElement = document.getElementById('input-autoPlay');

          self.properties.apiUrl = (<HTMLInputElement>apiUrlElement).value;
          self.properties.autoPlay = (<HTMLInputElement>autoPlayElement).checked;

          console.log(self.properties.apiUrl);
          console.log(self.properties.autoPlay);
        });

        setTimeout(async function () {
          console.log("Timeout expired. Running screenreading code.");          

          self.readPage(self);
        }, 3000);    
      }
    }
  }

  // Show an element
private show = function (elem) {
	elem.style.display = 'block';
};

// Hide an element
private hide = function (elem) {
	elem.style.display = 'none';
};

// Toggle element visibility
private toggle = function (elem) {

	// If the element is visible, hide it
	if (window.getComputedStyle(elem).display === 'block') {
		this.hide(elem);
		return;
	}

	// Otherwise, show it
	this.show(elem);

};

  private async readPage(aSelf)
  {
    if (aSelf.allText.length == 0)
    {
      aSelf.allText = aSelf.scrapePage();
    }        

    if (aSelf.allText.length > 0)
    {
      var aIndex = 1;
      let aSpeechResponse: Blob = await aSelf.atextToSpeechService.TextToSpeech(aSelf.allText[0]);

      var aObjectUrl: string = URL.createObjectURL(aSpeechResponse);
      var audio = new Audio();
      audio.src = aObjectUrl;
      audio.load();
      audio.play();

      audio.onended = async function()
      {
        if (aIndex < aSelf.allText.length)
        {
          let aSpeechResponse: Blob = await aSelf.atextToSpeechService.TextToSpeech(aSelf.allText[aIndex]);

          var aObjectUrl: string = URL.createObjectURL(aSpeechResponse);
          audio.src = aObjectUrl;
          // audio.load();
          audio.play();
          aIndex++;
        }
      };
    }
  }
  
  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
