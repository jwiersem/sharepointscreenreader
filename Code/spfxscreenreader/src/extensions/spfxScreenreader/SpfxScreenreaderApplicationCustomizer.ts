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

  private audio = new Audio();
  private currentAudioIndex:Number = 0;

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

  private scrapePage(): string[]
  {
    let allTextToRead: string[] = [];

    let heroElements: NodeListOf<Element> =  document.querySelectorAll(this.properties.selectors);

    console.log('Number of elements that might contain suitable text to read: ' + heroElements.length);

    for(var i = 0; i < heroElements.length; i++)
    {
      var aText: string = heroElements[i].getAttribute('aria-label');

      if (!aText)
      {
        aText = heroElements[i].textContent;
      }

      if (aText)
      {
        console.log(aText);
        allTextToRead.push(aText);
      }
      else{
        console.log('No text found for this element in aria-label or textContent properties.');        
      }
    }

    console.log("Number of suitable texts to play: " + allTextToRead.length);

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

    let config: IScreenreaderServiceConfiguration = {
      httpClient: this.context.httpClient,
      apiUrl: this.properties.apiUrl
    };

    this.atextToSpeechService = new ScreenreaderService(config);

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
          <div id="screenreader-audioplayer-wrapper" class="screenreader-audioplayer-wrapper">
            <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.top}">
              <h4>
                Screenreader controls
              </h4>
              <div class="icon-container">
                <i id="screenreader-rewind" class="ms-Icon ms-Icon--Rewind x-hidden-focus ms-fontColor-themeDarker--hover" title="Rewind" aria-hidden="true" data-reactid=".0.0.$=10.0.1.$=1$3.6.1.$0.0"></i>
              </div>
              
              <div class="icon-container">
                <i id="screenreader-stop" class="ms-Icon ms-Icon--CircleStopSolid x-hidden-focus ms-fontColor-themeDarker--hover" title="CircleStopSolid" aria-hidden="true" data-reactid=".0.0.$=10.0.1.$=1$3.6.1.$1.0"></i>
              </div>

              <div class="icon-container">
                <i id="screenreader-play" class="ms-Icon ms-Icon--Play x-hidden-focus ms-fontColor-themeDarker--hover" title="Play" aria-hidden="true" data-reactid=".0.0.$=10.0.1.$=1$3.6.1.$1.0"></i> 
              </div>   
            </div>
          </div>          
        </div>
        `;
        }     
           
        let self = this;

        document.getElementById('screenreader-rewind').addEventListener('click', function(){
          self.currentAudioIndex = 0;
          console.log('clicked rewind');
        });

        document.getElementById('screenreader-stop').addEventListener('click', function(){
          self.audio.pause();
          console.log('clicked stop');
        });

        document.getElementById('screenreader-play').addEventListener('click', function(){
          console.log('clicked play');

          if (!self.properties.autoPlay)
          {
            self.readPage(self);
          }
          else
          {
            if (self.audio.paused)
            {
              self.audio.play();
            }
            else
            {
              console.log('Auto play is on and audio is playing, cannot play audio twice.');
            }            
          }         
        });

        if (self.properties.autoPlay)
        {
          setTimeout(async function () {
            console.log("Timeout expired. Running screenreading code.");          

            self.readPage(self);
          }, 3000);  
        }  
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
    console.log('Start page reading');
    if (aSelf.allText.length == 0)
    {
      aSelf.allText = aSelf.scrapePage();
    }        

    if (aSelf.allText.length > 0)
    {
      let aSpeechResponse: Blob = await aSelf.atextToSpeechService.TextToSpeech(aSelf.allText[aSelf.currentAudioIndex]);

      var aObjectUrl: string = URL.createObjectURL(aSpeechResponse);
      
      aSelf.audio.src = aObjectUrl;
      aSelf.audio.load();
      aSelf.audio.play();

      aSelf.currentAudioIndex++;

      aSelf.audio.onended = async function()
      {
        if (aSelf.currentAudioIndex < aSelf.allText.length)
        {
          let aSpeechResponse: Blob = await aSelf.atextToSpeechService.TextToSpeech(aSelf.allText[aSelf.currentAudioIndex]);

          var aObjectUrl: string = URL.createObjectURL(aSpeechResponse);
          aSelf.audio.src = aObjectUrl;
          aSelf.audio.play();
          aSelf.currentAudioIndex++;
          console.log('Next audio clip number: ' + aSelf.currentAudioIndex);
        }
      };
    }
  }
  
  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
