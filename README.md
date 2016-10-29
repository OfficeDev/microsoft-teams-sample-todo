# Microsoft Teams - Todo MVC Sample

This sample demonstrates on how to quickly integrate your existing Web Application with **Microsoft Teams**.

The following sample has been modified from **TodoMVC for React** that can be found [here](https://github.com/tastejs/todomvc/tree/gh-pages/examples/typescript-react).

## Requirements

The sample requires `node.js` and uses the following stack:

1. [`React by Facebook`](https://facebook.github.io/react/) as the UI Framework - .
2. [`TypeScript`](https://www.typescriptlang.org/) as the transpiler.
4. [`TodoMVC`](http://todomvc.com/examples/typescript-react/#/) base for TodoMVC functionality.
5. [`Webpack`](https://webpack.github.io/) as the build tool.

## Get Started

1. Run `npm install`.
2. Run `npm start` to start the `webpack-dev-server`.
3. Alternatively, run `npm run build` to generate a deployable build.

### Authentication

To get Outlook tasks, you'll need to configure the authentication in the following manner:

1. Register your application at [Microsoft Application Registration Portal](https://apps.dev.microsoft.com/Landing?ru=https%3a%2f%2fapps.dev.microsoft.com%2f).
2. Click on `Add Platform` and choose `Web`.
3. Check `Allow Implicit Flow` and configure the redirect url to be either `https://localhost:3000/config.html` or `https://<mywebsite>/config.html`.
4. In the [`Outlook Tasks`](https://github.com/OfficeDev/microsoft-teams-sample-todo/blob/master/src/services/outlook.tasks.ts) file, edit line 22 as `clientId = <your clientId>`.

### Branches

While the `master` branch shows the latest state of the sample, you can also have a look at the following:

1. [`before`](https://github.com/OfficeDev/microsoft-teams-sample-todo/tree/before): shows the standard TodoMVC project with minor additions such as Authentication, OutlookTasks.

2. [`after`](https://github.com/OfficeDev/microsoft-teams-sample-todo/tree/before): shows the integration of the project in `before` with `Microsoft Teams`.

## Integrating with Microsoft Teams

In order to integrate with Microsoft Teams some specific changes need to be made. These changes are well showcased in this [code diff](https://github.com/OfficeDev/microsoft-teams-sample-todo/compare/before...after).

Walking through the changes the following can be observed:

1. We have added a new `config.html` & `config.tsx` page which is responsible for the the application to allow the user to manipulate any settings, perform authentication etc during the first launch.

2. We have added the same `config.html` file to our `webpack.common.js` configuration so that it can inject the right bundles during runtime.

3. We add a reference to `MicrosoftTeams.js` in our index.html & added `MicrosoftTeams.d.ts` for Typescript intellisense.

4. We have added a `manifest.json` and two logos for size *44x44* and *88x88*.

5. We have added some styles to our `app.css`.

6. Finally we have modified the `authentication` to depend on `Microsoft Teams` dialog instead.

## Microsoft Teams API

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