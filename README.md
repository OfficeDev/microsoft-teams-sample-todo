# Microsoft Teams 'Todo List' sample tab app

This is an example [tab app for Microsoft Teams](https://aka.ms/microsoftteamstabsplatform).  When it is added to a channel, it provides a basic task manager.
* Its purpose is to illustrate how simple it is to convert an existing web app into a Microsoft Teams tab app.  Take a look at the [code diff between the 'before' and 'after' branches](https://github.com/OfficeDev/microsoft-teams-sample-todo/compare/before...after).
* It is not a realistic example of a team collaboration app.  The tasks shown belong to the user's individual account and not to a shared team account. 

## Prerequisites

An [Office 365 account with access to Microsoft Teams](setup.md).

**TODO make this a link to setup.md in main docs repo**

## Run the app

### Register the app so it can access to Outlook tasks

1. Go to the [Microsoft Application Registration Portal](https://apps.dev.microsoft.com).
2. Sign in with your Office 365 work/school account.  Don't use your personal Microsoft account.
2. Add a new app.
2. Take note of your new `Application ID`.
2. Click on `Add Platform` and choose `Web`.
3. Check `Allow Implicit Flow` and configure the redirect URL to be either `https://localhost:3000/config.html` or `https://<mywebsite>/config.html`.

>**Note:** By defult, your organization should allow you to create new apps. But if it doesn't, you can use a one-year trial subscription of Office 365 Developer at no charge. [Here's how](setup.md).

**TODO make this a link to setup.md in main docs repo**

### Host the tab apps "configuration page" and "content page" 

1. Clone the repo.
4. Edit `src/services/outlook.tasks.ts` line 22 as `clientId = <your applicationId>`, using the application ID obtained above..
1. Install [node.js](https://nodejs.org) if you don't already have it. 
1. Run `npm install`.
2. Run `npm start` to start the `webpack-dev-server`.
3. Alternatively, run `npm run build` to generate a deployable build.

### Add the tab to Microsoft Teams

1. Download the [tab app package](https://github.com/OfficeDev/microsoft-teams-sample-todo/blob/master/package/ToDoTab.zip) for this sample. 
2. Create a new team for testing, if necessary.  Click **Create team** at the bottom of the left-hand panel.
3. Select the team from the left-hand panel, select **... (more options)** and then select **View Team**.
4. Select the **Developer (Preview)** tab, and then select **Upload**.
5. Navigate to the zip file and select it.
6. Go to any channel in the team.  Click the '+' to the right of the existing tabs. 
7. Select your tab from the gallery that appears.
8. Accept the consent prompt.
9. Sign in using your Office 365 work/school account, and click Save.

> **Note:** To re-upload an updated package, with the same `id`, click the 'Replace' icon at the end of the tab's table row.  Don't click 'Upload' again: Microsoft Teams will say the tab already exists.

## Code walk through

While the `master` branch shows the latest state of the sample, take a look at the following [code diff](https://github.com/OfficeDev/microsoft-teams-sample-todo/compare/before...after)between:

* [`before`](https://github.com/OfficeDev/microsoft-teams-sample-todo/tree/before): the initial app

* [`after`](https://github.com/OfficeDev/microsoft-teams-sample-todo/tree/after): the app after integration with Microsoft Teams.

Going through this step by step:

1. We have added a new `config.html` & `config.tsx` page which is responsible for the the application to allow the user to manipulate any settings, perform authentication etc during the first launch.

2. We have added the same `config.html` file to our `webpack.common.js` configuration so that it can inject the right bundles during runtime.

3. We add a reference to `MicrosoftTeams.js` in our index.html & added `MicrosoftTeams.d.ts` for Typescript intellisense.

4. We have added a `manifest.json` and two logos for size *44x44* and *88x88*.

5. We have added some styles to our `app.css`.

6. Finally we have modified the `authentication` to depend on `Microsoft Teams` dialog instead.

### Handling the 'Save' event

```typescript
completeSave() {
    if (!microsoftTeams) return;
    microsoftTeams.settings.setValidityState(true);
    microsoftTeams.settings.registerOnSaveHandler(saveEvent => {
        microsoftTeams.settings.setSettings({
            contentUrl: `${location.origin}/index.html`,
            suggestedDisplayName: "My Tasks",
            websiteUrl: `${location.origin}/index.html`
        });
        saveEvent.notifySuccess();
    });
}
```

### Invoking the Authentication dialog

```typescript
microsoftTeams.authentication.authenticate({
    url: params.url,
    width: windowSize.toPixels().width,
    height: windowSize.toPixels().height,
    successCallback: successCallback,
    failureCallback: exception => {
        return reject(new OAuthError('Error while launching dialog: ' + JSON.stringify(exception)));
    }
});
```

## Technology used

The following sample has been modified from **TodoMVC for React** that can be found [here](https://github.com/tastejs/todomvc/tree/gh-pages/examples/typescript-react).

It uses the following stack:

1. [`React by Facebook`](https://facebook.github.io/react/) as the UI Framework - .
2. [`TypeScript`](https://www.typescriptlang.org/) as the transpiler.
4. [`TodoMVC`](http://todomvc.com/examples/typescript-react/#/) base for TodoMVC functionality.
5. [`Webpack`](https://webpack.github.io/) as the build tool.

