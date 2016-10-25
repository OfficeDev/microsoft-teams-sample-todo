///<reference path="./MicrosoftTeams.d.ts"/>

((microsoftTeams) => {
    "use strict";

    if (!microsoftTeams) return;

    microsoftTeams.initialize();
    microsoftTeams.settings.setValidityState(true);
    microsoftTeams.settings.registerOnSaveHandler(saveEvent => {
        microsoftTeams.settings.setSettings({ contentUrl: "/index.html", suggestedTabName: "React • TodoMVC", websiteUrl: "https://localhost:3000" });
        saveEvent.notifySuccess();
    });

})((window as any).microsoftTeams);