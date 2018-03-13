import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

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
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpfxScreenreaderApplicationCustomizer
  extends BaseApplicationCustomizer<ISpfxScreenreaderApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;
  private atextToSpeechService: ScreenreaderService;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let config: IScreenreaderServiceConfiguration = {
      httpClient: this.context.httpClient,
      apiUrl: "https://prod-32.westeurope.logic.azure.com:443/workflows/737b64d81a9e4dc8b0dd1b938789df2b/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=PQbLDWnyX2m7GK6wRF5p03dvPacpYYPh87h5jp352dM"
    };

    this.atextToSpeechService = new ScreenreaderService(config);

    // Added to handle possible changes on the existence of placeholders.
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    // Call render method for generating the HTML elements.
    this._renderPlaceHolders();
    return Promise.resolve<void>();
  }

  private async _renderPlaceHolders(): Promise<void> {

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

        let aSpeechResponse: Blob = await this.atextToSpeechService.TextToSpeech("First line from application customizer.");

        var audio = new Audio();
        audio.src = URL.createObjectURL(aSpeechResponse);
        audio.load();
        audio.play();

        if (this._topPlaceholder.domElement) {
          this._topPlaceholder.domElement.innerHTML = `
        <div class="${styles.app}">
          <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.top}">
            <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${ audio.outerHTML }
          </div>
        </div>
        `;
        }        
      }
    }
  }

  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
