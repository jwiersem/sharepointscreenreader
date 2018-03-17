## spfxscreenreader

This is hobby project that will allow you to hear your SharePoint page speak to you!
When installed the code will read the most important text on the page aloud.
It is an SharePoint Framework application customizer which means that when it is installed it will run on every page in your SharePoint site collection.
At the top of the page you will see an audio player, with rewind, stop and play buttons.
Installing the package will also deploy a list to your site, in which you will need to specify the following settings:
* API Url 
* Autoplay (yes/no)
* The selectors you wish to use for finding the important elements on the page.
This repository contains the app package that needs to be installed in your SharePoint farm. It also contains the definition of the API the code requires. It is implemented as a Microsoft Flow that is HTTP triggered and returns a HTTP response. The Microsoft Flow in turn uses the Bing Speech API that generates audio from text.
The app package and the Flow definition are in the release folder.

These are the installation instructions:


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
