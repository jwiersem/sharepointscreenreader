import { HttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';

import { ScreenreaderService, IScreenreaderServiceConfiguration } from '../Services/ScreenreaderService';
import { ISpfxScreenreaderApplicationCustomizerProperties } from '../SpfxScreenreaderApplicationCustomizer';
/**
 * @class
 * Screenreader control
 * */
export class Screenreader {

    private audio = new Audio();
    private currentAudioIndex: number = 0;

    private atextToSpeechService: ScreenreaderService;
    private allText: string[] = [];

    private httpClient: HttpClient;

    private properties: ISpfxScreenreaderApplicationCustomizerProperties;

    public constructor(aHttpClient: HttpClient, aProperties: ISpfxScreenreaderApplicationCustomizerProperties) {
        this.properties = aProperties;
        this.httpClient = aHttpClient;

        let config: IScreenreaderServiceConfiguration = {
            httpClient: this.httpClient,
            apiUrl: this.properties.apiUrl
        };

        this.atextToSpeechService = new ScreenreaderService(config);
    }

    public render(styles): string {
        return `
        <div class="${styles.app}">
          <div id="screenreader-audioplayer-wrapper" class="screenreader-audioplayer-wrapper">
            <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.top}">
                <div class="${styles.screenreadercontrols}">
                    <h4>
                        Screenreader controls
                    </h4>
                    <div class="${styles.iconcontainer}">
                        <i id="screenreader-rewind" class="ms-Icon ms-Icon--Rewind x-hidden-focus ms-fontColor-themeDarker--hover" title="Rewind" aria-hidden="true" data-reactid=".0.0.$=10.0.1.$=1$3.6.1.$0.0"></i>
                    </div>
                    
                    <div class="${styles.iconcontainer}">
                        <i id="screenreader-stop" class="ms-Icon ms-Icon--CircleStopSolid x-hidden-focus ms-fontColor-themeDarker--hover" title="Stop" aria-hidden="true" data-reactid=".0.0.$=10.0.1.$=1$3.6.1.$1.0"></i>
                    </div>

                    <div class="${styles.iconcontainer}">
                        <i id="screenreader-play" class="ms-Icon ms-Icon--Play x-hidden-focus ms-fontColor-themeDarker--hover" title="Play" aria-hidden="true" data-reactid=".0.0.$=10.0.1.$=1$3.6.1.$1.0"></i> 
                    </div>   
                </div>
            </div>
          </div>          
        </div>
        `;
    }

    public addInteractivity() {
        document.getElementById('screenreader-rewind').addEventListener('click', (event) => {
            this.currentAudioIndex = 0;
            console.log('clicked rewind');
        });

        document.getElementById('screenreader-stop').addEventListener('click', (event) => {
            this.audio.pause();
            console.log('clicked stop');
        });

        document.getElementById('screenreader-play').addEventListener('click', (event) => {
            console.log('clicked play');

            if (!this.properties.autoPlay) {
                this.readPage();
            }
            else {
                if (this.audio.paused) {
                    this.audio.play();
                }
                else {
                    console.log('Auto play is on and audio is playing, cannot play audio twice.');
                }
            }
        });

        if (this.properties.autoPlay) {                     
            setTimeout(async (event) => {
                console.log("Running screenreading code.");

                this.readPage();
            }, 5000);
        }
    }

    // Show an element
    private show(elem): void {
        elem.style.display = 'block';
    }

    // Hide an element
    private hide(elem): void {
        elem.style.display = 'none';
    }

    // Toggle element visibility
    private toggle(elem): void {

        // If the element is visible, hide it
        if (window.getComputedStyle(elem).display === 'block') {
            this.hide(elem);
            return;
        }

        // Otherwise, show it
        this.show(elem);

    }


    private scrapePage(): string[] {
        let allTextToRead: string[] = [];

        let heroElements: NodeListOf<Element> = document.querySelectorAll(this.properties.selectors);

        console.log('Number of elements that might contain suitable text to read: ' + heroElements.length);

        for (var i = 0; i < heroElements.length; i++) {
            var aText: string = heroElements[i].getAttribute('aria-label');

            if (!aText) {
                aText = heroElements[i].textContent;
            }

            if (aText) {
                console.log(aText);
                allTextToRead.push(aText);
            }
            else {
                console.log('No text found for this element in aria-label or textContent properties.');
            }
        }

        console.log("Number of suitable texts to play: " + allTextToRead.length);

        return allTextToRead;
    }

    private async readPage() {
        console.log('Start page reading');
        if (this.allText.length == 0) {
            this.allText = this.scrapePage();
        }

        if (this.allText.length > 0) {
            let aSpeechResponse: Blob = await this.atextToSpeechService.TextToSpeech(this.allText[this.currentAudioIndex]);

            var aObjectUrl: string = URL.createObjectURL(aSpeechResponse);

            this.audio.src = aObjectUrl;
            this.audio.load();
            this.audio.play();

            this.currentAudioIndex++;

            this.audio.onended = async (event) => {
                if (this.currentAudioIndex < this.allText.length) {
                    let aSpeechResponse: Blob = await this.atextToSpeechService.TextToSpeech(this.allText[this.currentAudioIndex]);

                    var aObjectUrl: string = URL.createObjectURL(aSpeechResponse);
                    this.audio.src = aObjectUrl;
                    this.audio.play();
                    this.currentAudioIndex++;
                    console.log('Next audio clip number: ' + this.currentAudioIndex);
                }
            };
        }
    }
}