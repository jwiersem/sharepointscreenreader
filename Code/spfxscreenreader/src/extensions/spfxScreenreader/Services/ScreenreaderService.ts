import { HttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
// import JsonUtilities from '@microsoft/sp-core-library';

/**
 * @interface
 * Interface for ScreenreaderService configuration
 */
export interface IScreenreaderServiceConfiguration {
    httpClient: HttpClient;
    apiUrl: string;
  }

  /**
 * @interface
 * Interface for body data 
 */
export interface ISpeechAPIBodyData {
    language: string;
    gender: string;
    text:string;
  }

/**
 * @class
 * Service to do text to speech
 * */
export class ScreenreaderService
{
    private httpClient: HttpClient;
    private apiUrl: string;

    constructor(config: IScreenreaderServiceConfiguration){
        this.httpClient = config.httpClient;
        this.apiUrl = config.apiUrl;
    }

    public TextToSpeech(aText:string) : Promise<Blob>
    {
        let aBody: ISpeechAPIBodyData = {
            language:"en-us",
            gender:"Female",
            text: aText
        };

        let httpPostOptions: IHttpClientOptions = {
            headers: {
              "content-type": "application/json"
            },
            body: JSON.stringify(aBody)
          };        
        
        return this.httpClient.post(this.apiUrl, HttpClient.configurations.v1, httpPostOptions).then((response: HttpClientResponse) => {
            if (response.ok) {
                console.log("Returned OK from httpClient");
               return response.blob();
            } else {
                console.log("WARNING - failed to hit URL " + this.apiUrl + ". Error = " + response.statusText);
            }
          });
    }

}