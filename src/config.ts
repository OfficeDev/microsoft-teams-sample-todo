///<reference path="./MicrosoftTeams.d.ts"/>
declare var OfficeHelpers: any;

(() => {
    "use strict";

    if (!microsoftTeams) return;
    microsoftTeams.initialize();

    var authenticator = new OfficeHelpers.Authenticator();
    authenticator.endpoints.registerMicrosoftAuth('73d044ea-4ae0-4bd8-b26c-7c2f924410a2', {
        scope: 'https://outlook.office.com/tasks.readwrite'
    });

    authenticator.authenticate('Microsoft')
        .then(() => microsoftTeams.settings.setValidityState(true));

    microsoftTeams.settings.registerOnSaveHandler(saveEvent => {
        microsoftTeams.settings.setSettings({
            contentUrl: "https://localhost:3000/index.html",
            suggestedDisplayName: "My Tasks",
            websiteUrl: "https://localhost:3000/index.html"
        });
        saveEvent.notifySuccess();
    });
})();