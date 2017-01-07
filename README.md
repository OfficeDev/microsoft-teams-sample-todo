# Microsoft Teams 'Todo List' sample tab app

This is an example [tab app for Microsoft Teams](https://aka.ms/microsoftteamstabsplatform).  The point of this sample to illustrate how simple it is to convert an existing web app into a Microsoft Teams tab app.  The existing web app, [**TodoMVC for React**](https://github.com/tastejs/todomvc/tree/gh-pages/examples/typescript-react), provides a basic task manager which integrates with your personal Outlook Tasks. With only a few minor modifications, this web view can be added to a channel as a tab app.  Take a look at the [code diff between the 'before' and 'after' branches](https://github.com/OfficeDev/microsoft-teams-sample-todo/compare/85ac809a2b52b528e8323a0a14e419afca21da12...9a1224eb276fa15a76f9e4882c4abe5ae8b68a99) to see what changes were made.

> **Note:** This is not a realistic example of a team collaboration app.  The tasks shown belong to the user's individual account and not to a shared team account.

**For more information on developing experiences for Microsoft Teams, please review the Microsoft Teams [developer documentation](https://msdn.microsoft.com/en-us/microsoft-teams/index).**

## Prerequisites

1. An [Office 365 account with access to Microsoft Teams](https://msdn.microsoft.com/en-us/microsoft-teams/setup).
2. This sample is built using [Node.js](https://nodejs.org).  Download and install the recommended version if you don't already have it.

## Run the app

### Host the tab apps "configuration page" and "content page"

To enable the app in Dev mode:
1. Clone the repo.
2. Open a command line in the repo subdirectory.
3. Run `npm install` (included as part of Node.js) from the command line.
4. Run `npm start` to start the `webpack-dev-server` to enable the dev app to function.
5. Alternatively, run `npm run build` to generate a deployable build, which you can host in your own environment.  Note:  you will need to modify the config.url in the manifest to point to your hosting location. [More information](#registering-an-application-to-authenticate-with-microsoft)

### Add the tab to Microsoft Teams

1. Download the [tab app dev package](https://github.com/OfficeDev/microsoft-teams-sample-todo/raw/before-final/package/todo.dev.zip) zip file for this sample.
2. Create a new team for testing, if necessary. Click **Create team** at the bottom of the left-hand panel.
3. Select the team from the left-hand panel, select **... (more options)** and then select **View Team**.
4. Select the **Developer (Preview)** tab, and then select **Upload**.
5. Navigate to the downloaded zip file from step 1 above and select it.
6. Go to any channel in the team.  Click the '+' to the right of the existing tabs.
7. Select your tab from the gallery that appears.
8. Accept the consent prompt.
9. If needed, sign in using your Office 365 work/school account.  Note that the code will try to do silent authentication if possible.
10. Validate authentication information.
11. Hit Save to add the tab to channel.

> **Note:** To re-upload an updated package, with the same `id`, click the 'Replace' icon at the end of the tab's table row.  Don't click 'Upload' again: Microsoft Teams will say the tab already exists.

> It is advisable to have multiple configs, one per environment. The names of the zip files can be anything such as `todo.dev.zip`, `todo.prod.zip` etc. but the zip must contain a `manifest.json` with a unique `id`.  See [Creating a manifest for your tab](https://msdn.microsoft.com/en-us/microsoft-teams/createpackage).

## Code walk through

While the `master` branch shows the latest state of the sample, take a look at the following [code diff](https://github.com/OfficeDev/microsoft-teams-sample-todo/compare/85ac809a2b52b528e8323a0a14e419afca21da12...9a1224eb276fa15a76f9e4882c4abe5ae8b68a99) between:

* [`before`](https://github.com/OfficeDev/microsoft-teams-sample-todo/commit/85ac809a2b52b528e8323a0a14e419afca21da12): the initial app

* [`after`](https://github.com/OfficeDev/microsoft-teams-sample-todo/commit/9a1224eb276fa15a76f9e4882c4abe5ae8b68a99): the app after integration with Microsoft Teams.

Going through this step by step:

1. We have added a new `config.html` and `config.tsx` page which is responsible for the the application to allow the user to manipulate any settings, perform single signon Authentication etc. during the first launch. This is required so that the team administrator can configure the application/settings.  See [Create the configuration page](https://msdn.microsoft.com/en-us/microsoft-teams/createconfigpage).

2. We have added the same `config.html` file to our `webpack.common.js` configuration so that it can inject the right bundles during runtime.

3. We add a reference to `MicrosoftTeams.js` in our index.html and added `MicrosoftTeams.d.ts` for Typescript intellisense.

4. We have added a `manifest.dev.json`, `manifest.prod.json` and two logos for size *44x44* and *88x88*. Remember to rename these as `manifest.json` in your zip files that you upload to Microsoft Teams.

5. We have added some styles to our `app.css`.

6. Finally we have modified the `authentication` in outlook.tasks.ts to depend instead on 'useMicrosoftTeamsAuth', a new feature from the beta version of OfficeHelpers referenced in this sample.

### Invoking the Authentication dialog

When the user adds the tab, the configuration page is presented (config.html).  In this case, the code authenticates the user if possible.

Authentication leverages a new Teams-specific function in the latest (>0.4.0) version of the [office-js-helpers](https://github.com/OfficeDev/office-js-helpers) library.  This helper function will attempt to silently authenticate, but if it cannot, it will call the Microsoft Teams specific auth dialog for you.  For more information on the full Authentication process in Microsoft Teams, please review [Authenticating a user](https://msdn.microsoft.com/en-us/microsoft-teams/auth) in the Microsoft Teams [developer documentation](https://msdn.microsoft.com/en-us/microsoft-teams/index).

In outlook.tasks.ts:
```typescript
 return this.authenticator.authenticate('Microsoft')
    .then(token => this._token = token)
    .catch(error => {
        Utilities.log(error);
        throw new Error('Failed to login using your Microsoft Account');
    });
```

### Handling the 'Save' event

After successful sign-in, the user will Save the tab into the channel.  The following code enables the Save button, and sets the SaveHandler, which will store what content to display in the tab (in this case, just the project's index.html).

In config.tsx:
```typescript
    initialize({ groupId, upn}) {
        this.setState({ groupId, upn });
        console.log(this.state);
        /** Enable the Save button  */
        microsoftTeams.settings.setValidityState(true);
        /** Register the save handler */
        microsoftTeams.settings.registerOnSaveHandler(saveEvent => {
            /** Store Tab content settings */
            microsoftTeams.settings.setSettings({
                contentUrl: `${location.origin}/index.html`,
                suggestedDisplayName: "My Tasks",
                websiteUrl: `${location.origin}/index.html`
            });
            saveEvent.notifySuccess();
        });
    }
```

## Technology used


It uses the following stack:

1. [`React by Facebook`](https://facebook.github.io/react/) as the UI Framework.
2. [`TypeScript`](https://www.typescriptlang.org/) as the transpiler.
4. [`TodoMVC`](http://todomvc.com/examples/typescript-react/#/) base for TodoMVC functionality.
5. [`Webpack`](https://webpack.github.io/) as the build tool.

## Additional Resources

### Registering an application to authenticate with Microsoft
Note that this will not be necessary if you use the local Dev option above, but if you choose to host this tab in your own environment, you must register the application in order to authenticate.

1. Go to the [Microsoft Application Registration Portal](https://apps.dev.microsoft.com).
2. Sign in with your Office 365 work/school account.  Don't use your personal Microsoft account.
2. Add a new app.
2. Take note of your new `Application ID`.
2. Click on `Add Platform` and choose `Web`.
3. Check `Allow Implicit Flow` and configure the redirect URL to be `https://<mywebsite>/config.html`.

For more information on hosting your own tab pages, see the [Microsoft Teams 'Get Started' sample README](https://github.com/OfficeDev/microsoft-teams-sample-get-started#host-tab-pages-over-https).

>**Note:** By defult, your organization should allow you to create new apps. But if it doesn't, you can use a one-year trial subscription of Office 365 Developer at no charge. [Here's how](https://msdn.microsoft.com/en-us/microsoft-teams/setup).


## Credits

This project is based on the TodoMVC Typescript - React template located [here](https://github.com/tastejs/todomvc/tree/gh-pages/examples/typescript-react).

## Contributing

Please read [Contributing](contributing.md) for details on our code of conduct, and the process for submitting pull requests to us.

## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/officedev/microsoft-teams-sample-todo/tags).

## License

This project is licensed under the MIT License - see the [License](LICENSE) file for details