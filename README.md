## spfxscreenreader

**This is hobby project that will allow you to hear SharePoint speak to you!**

[![Youtube video for SPFx screenreader](https://www.youtube.com/upload_thumbnail?v=rE9CzuWchCg&t=hqdefault&ts=1521307739542)](https://www.youtube.com/watch?v=rE9CzuWchCg)

When installed the code will read the most important text on the page aloud.
It is a SharePoint Framework application customizer which means that will run on every page in your SharePoint site collection.
At the top of the page you will see an audio player, with rewind, stop and play buttons.
Installing the package will also deploy a list to your site, in which you will need to specify the following settings:
* API Url 
* Autoplay (yes/no)
* The selectors you wish to use for finding the important elements on the page.

This repository contains the app package that needs to be installed in your SharePoint farm. It also contains the definition of the API the code requires. It is implemented as a Microsoft Flow that is HTTP triggered and returns a HTTP response. The Microsoft Flow in turn uses the Bing Speech API that generates audio from text.
*The app package and the Flow definition are in the release folder.*
Like I said, this is a hobby/demo project. It is not suited to be used anywhere near a production environment. The main purpose was to show that besides the, obviously extremely useful scenario of seeing whether people are pissed off at you in their emails with sentiment analysis, there are also other ways to use Cognitive Services for the betterment of mankind.

These are the installation instructions:
1. You need programmatic access to the text-to-speech service I used, which is the Bing Speech API.
Register for a key here: [Register for Bing Speech API](https://azure.microsoft.com/en-us/try/cognitive-services/?api=speech-api "Register for Bing Speech API")
Save the key somehere.
2. Use the Flow/Logic App definition to create the 'backend'. If you have a personal O365 account you cant actually easily import the Flow :-( In that case you have the option to read JSON and recreate the Flow yourself, or if you have an Azure account you can create a Logic App from the JSON definition.
3. Fill in the Bing Speech API key in the apiKey variable in the Flow/LogicApp.
4. Enable the Flow/Logic App.
5. Now, edit the Flow again. You will see a URL in the trigger, save that somewhere.
6. Go to SharePoint (online) and to your App catalog site. Upload the sppkg file there in the Apps for SharePoint section. Accept the message because you like to live dangerously.
7. Go to a Modern type site, or create one. I advise a Communication site.
8. Click add App in the settings menu and add the app: 
spfxscreenreader-client-side-solution
9. After a while the list *Screenreadersettings* will be created in your top site in the site collection.
10. Open it and save a single new item with the following info:
* apiUrl: Fill in the URL that was magically created for you in the Flow.
* Enable Autoplay because it can in no way be annoying that every page in your SharePoint site will start to speak to you.
* Fill in css selectors that identity the HTML that have the text you want read. For example, fill in: *.ms-FocusZone, h1, h2, h3*
You can play with these selectors to get the App to read the text you want. You would need to read the HTML code of the pages to find out the right selectors. The App will read the aria-label or textContent properties of the elements you identify with the selectors.
11. Enable your sound..
12. Go to the homepage of your communication site and enjoy hearing SharePoint speak to you, like it has to me for so many years..

The good:
* Very simple demo project to show options for increasing accessibility of SharePoint.
* Shows how easy it is to use cognitive services.
* Use a Flow/Logic App as an API that you can edit visually.
* Nice for trying out SPFx Application customizers.
* Ample opportunity to criticize the code, visual design, intelligence and character of the author.
* Free of charge!

The bad:
* Very quick and dirty solution.
* It turns out you need an organisation Flow license to import-export Flows. This completely annihilated my plan to make the installation consist of 100% clicking and copy-pasting.
* Will not actually help improve the current dismal state of mankind in general, or your own life in particular.


### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean
gulp serve 
gulp bundle 
gulp package-solution
