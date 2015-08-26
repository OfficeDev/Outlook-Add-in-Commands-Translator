# Outlook-Add-in-Commands-Translator

The Translator add-in uses the commands model for Outlook add-ins to add a button to the ribbon in the new message form. The button sends the selected text from the message body to a translator web service to translate from English to Russian.

## Prerequsites

In order to run this sample, you will need the following:

- A web server to host the sample files. The server must be able to accept SSL-protected requests (https) and have a valid SSL certificate.
- An Office 365 email account **or** an Outlook.com email account.
- An API key for the [Yandex Translate API](https://translate.yandex.com/developers).
- Outlook 2016, which is part of the [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview).

## Configuring and installing the sample

1. Download or fork the repository.
1. Open `TranslateHelper.js` in a text editor and replace the `YOUR API KEY HERE` value with your API key for the Yandex Translate API. Save your changes.
1. Upload the `AppCompose`, `Content`, and `Images` directories to a directory on your web server.
1. Open `TranslateAppManifest.xml` in a text editor. Replace all instances of `YOUR_WEB_SERVER` with the HTTPS URL of the directory where you uploaded the files in the previous step. Save your changes.
1. Logon to your email account with a browser at either https://outlook.office365.com (for Office 365), or https://www.outlook.com (for Outlook.com). Click on the gear icon in the upper-right corner, then click **Manage add-ins**.
    
  ![The Manage add-ins menu item on https://www.outlook.com](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/outlook-manage-addins.PNG)
    
1. In the add-in list, click the **+** icon and choose **Add from a file**.

  ![The Add from file menu item in the add-in list](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/addin-list.PNG)

1. Click **Browse** and browse to the `TranslateAppManifest.xml` file on your development machine. Click **Next**.

  ![The Add add-in from a file dialog](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/browse-manifest.PNG)

1. On the confirmation screen, you will see a warning that the add-in is not from the Office Store and hasn't been verified by Microsoft. Click **Install**.
1. You should see a success message: **You've added an add-in for Outlook**. Click OK.

## Running the sample ##

1. Open Outlook 2016 and connect to the email account where you installed the add-in.
1. Create a new email. Notice that the add-in has placed a **Translate** button on the command ribbon.

  ![The Translate button on a new mail form in Outlook](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/new-mail.PNG)

1. Type some English text into the body. Select the text, then click **Translate**.

  ![The new mail form with English text selected in the body](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/text-selected.PNG)

1. After a moment, the selected text will be replaced with the Russian translation, and you should see the message **Translated successfully** in the information bar.

  ![The text translated into Russian by the add-in](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/text-translated.PNG)

## Key components of the sample

- `TranslateAppManifest.xml`: The manifest file for the Translator add-in.
- `AppCompose/FunctionFile/Home.html`: An empty HTML file to load `Translator.js` for clients that support add-in commands.
- `AppCompose/FunctionFile/Translator.js`: The code that is invoked when the add-in command button is clicked.
- `AppCompose/Home/Home.html`: The HTML file that is loaded and displayed by clients that do not support add-in commands.
- `AppCompose/Home/Home.js`: The code that is invoked by clients that do not support add-in commands.
- `TranslateHelper.js`: Common code used by both `Translator.js` and `Home.js`.

## How's it all work?

The key part of the sample is the structure of the manifest file. The manifest uses the same version 1.1 schema as any Office add-in's manifest. However, there is a new section of the manifest called `VersionOverrides`. This section holds all of the information that clients that support add-in commands (currently only Outlook 2016) need to invoke the add-in from a ribbon button. By putting this in a completely separate section, the manifest can also include the original markup to enable the add-in to be loaded in a task pane under the original compose mode add-in model. You can see this in action by loading the add-in in Outlook 2013 or Outlook on the web.

### The Translator add-in loaded in Outlook on the web ###

![The Translator add-in loaded in Outlook on the web](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/outlook-on-web.PNG)

Within the `VersionOverrides` element, there are two child elements, `Resources` and `Hosts`. The `Resources` element contains information about icons, strings, and what HTML file to load for the add-in. The `Hosts` section specifies how and when the add-in is loaded.

In this sample, there is only one host specified (Outlook):

    <Host xsi:type="MailHost">
    
Within this element are the configuration specifics for the desktop version of Outlook:

    <DesktopFormFactor>
    
The URL to the HTML file with all of the JavaScript code for the button is specified in the `FunctionFile` element (note that it uses the resource ID specified in the `Resources` element):

    <FunctionFile resid="functionFile" />
    
The manifest also limits activation to the new message form by setting a single extension point:

    <ExtensionPoint xsi:type="MessageComposeCommandSurface">
    
The properties of the button are specified in the `Control` element. Most importantly, the button's click event is connected to the `translate` function in `Translator.js` inside the `Action` element:

    <Action xsi:type="ExecuteFunction">
      <FunctionName>translate</FunctionName>
    </Action>
    
## Questions and comments

- If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/Outlook-Add-in-Commands-Translator/issues).
- Questions about Office Add-in development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Make sure that your questions or comments are tagged with `office-addins`.

## Additional resources

- [Outlook Dev Center](https://dev.outlook.com)
- [Office Add-ins](https://msdn.microsoft.com/library/office/jj220060.aspx) documentation on MSDN
- [More Add-in samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)

## Copyright

Copyright (c) 2015 Microsoft. All rights reserved.