---
page_type: sample
products:
- office-teams
- office-outlook
- office-365
languages:
- typescript
- nodejs
extensions:
  contentType: samples
  technologies:
  - Tabs
  createdDate: 10/19/2016 12:06:51 PM
---
# Exemple d’application Tab « Liste de tâches » pour Microsoft Teams
[![État de création](https://travis-ci.org/OfficeDev/microsoft-teams-sample-todo.svg?branch=master)](https://travis-ci.org/OfficeDev/microsoft-teams-sample-todo)

Il s’agit d’un exemple d’[application Tab pour Microsoft Teams](https://aka.ms/microsoftteamstabsplatform). Cet exemple illustre à quel point il est simple de convertir une application web existante en une application Tab pour Microsoft Teams. L’application web existante, [**TodoMVC pour React**](https://github.com/tastejs/todomvc/tree/gh-pages/examples/typescript-react), fournit un gestionnaire des tâches de base qui s’intègre à vos tâches Outlook personnelles. Avec seulement quelques modifications mineures, vous pouvez ajouter cet affichage web à un canal sous la forme d’une application à onglets. Examinez le [code diffère entre les branches « Avant » et « après »](https://github.com/OfficeDev/microsoft-teams-sample-todo/compare/before...after) pour voir quelles modifications ont été apportées.

> **Remarque :** Il ne s’agit pas d’un exemple réaliste d’une application de collaboration d’équipe. Les tâches affichées appartiennent au compte individuel de l'utilisateur et non à un compte d'équipe partagé.

**Pour plus d’informations sur le développement d’expériences pour Microsoft Teams, veuillez consulter la [documentation des développeurs](https://msdn.microsoft.com/en-us/microsoft-teams/index) de Microsoft Teams.**

## Conditions préalables

1. Un [compte Office 365 avec accès à Microsoft Teams](https://msdn.microsoft.com/en-us/microsoft-teams/setup).
2. Cet exemple est créé à l’aide de [Node.js](https://nodejs.org). Téléchargez et installez la version recommandée si vous ne l’avez pas encore.

## Exécuter l’application

### Héberger les applications onglet « page de configuration » et « page de contenu »

Lorsque la configuration de l’onglet et les pages de contenu sont des applications web, celles-ci doivent être hébergées.  

Tout d’abord, nous allons préparer le projet :

1. Clonez le référentiel.
2. Ouvrez une ligne de commande dans le sous-répertoire référentiel.
3. Exécutez `npm install` (inclus dans le cadre de node.js) à partir de la ligne de commande.

Pour héberger l’application en local :
* dans le sous-répertoire référentiel, exécutez `npm start` pour démarrer le `webpack-dev-server` pour activer l’application de développement.
* Notez que le manifeste dev est configuré pour utiliser localhost. Aucune configuration supplémentaire n’est nécessaire.

Si vous le souhaitez, vous pouvez créer une application déployable que vous pouvez héberger dans votre propre environnement :
* dans le sous-répertoire référentiel, exécutez `npm run build` pour générer une build pouvant être déployée.
* Modifiez l'adresse config.url dans le manifeste de l'application pour qu'elle pointe vers votre site d'hébergement. [Voir ci-dessous](#registering-an-application-to-authenticate-with-microsoft)

### Ajouter l’onglet à Microsoft Teams

1. Téléchargez le fichier zip [package de développement d’applications de l’onglet](https://github.com/OfficeDev/microsoft-teams-sample-todo/tree/master/package/todo.dev.zip) pour cet exemple.
2. Créez une équipe à des fins de test, si nécessaire. Dans la partie inférieure du volet gauche, cliquez sur **Créer une équipe**.
3. Sélectionnez l’équipe dans le volet gauche, sélectionnez **... (autres options)** puis sélectionnez **Afficher l’équipe**.
4. Sélectionnez l’onglet **Développeur (Aperçu)**, puis sélectionnez **Télécharger**.
5. Accédez au fichier zip téléchargé à l’étape 1 ci-dessus et sélectionnez-le.
6. Accédez à un canal quelconque de l’équipe. Cliquez sur le signe « + » à droite des onglets existants.
7. Sélectionnez l’onglet de la galerie qui s’affiche.
8. Acceptez l’invite de consentement.
9. Si nécessaire, connectez-vous à l’aide de votre compte professionnel/scolaire Office 365. Veuillez noter que le code essaiera d’effectuer une authentification silencieuse si possible.
10. Validez les informations d’authentification.
11. Cliquez sur Enregistrer pour ajouter l’onglet au canal.

> **Remarque :** Pour retélécharger un package mis à jour, avec le même de `ID`, cliquez sur l’icône de remplacement à la fin de la ligne de tableau de l’onglet. Ne cliquez pas de nouveau sur « Télécharger » : Microsoft Teams indique que l’onglet existe déjà.

> Il est recommandé de disposer de plusieurs configurations, une par environnement. Les noms des fichiers zip peuvent être tels que `todo.dev.zip`, `todo.prod.zip` etc., mais le zip doit contenir un `manifest.json` avec un `ID`unique. Consultez [Création d’un manifeste de tabulation](https://msdn.microsoft.com/en-us/microsoft-teams/createpackage).

## Guide de code

Bien que la branche `master` présente l’état le plus récent de l’exemple, examinez le [code diff](https://github.com/OfficeDev/microsoft-teams-sample-todo/compare/before...after) suivant entre :

* [`avant`](https://github.com/OfficeDev/microsoft-teams-sample-todo/commit/before): l’application initiale

* [`après`](https://github.com/OfficeDev/microsoft-teams-sample-todo/commit/after) : l’application après intégration avec Microsoft Teams.

Procédez comme suit pour passer à l’étape suivante :

1. Nous avons ajouté une nouvelle page `config.html` et `config.tsx` qui est responsable de l'application pour permettre à l'utilisateur de manipuler tous les paramètres, d'effectuer l'authentification par signature unique, etc. lors du premier lancement. Celle-ci est nécessaire pour permettre à l’administrateur de l’équipe de configurer l’application/les paramètres. Consultez [Créer la page de configuration](https://msdn.microsoft.com/en-us/microsoft-teams/createconfigpage).

2. Nous avons ajouté le même fichier `config.html` à la configuration `webpack.common.js` de sorte qu’il puisse injecter les offres groupées appropriées pendant l’exécution.

3. Nous ajoutons une référence à `MicrosoftTeams. js` dans notre index.html et ajoutons `MicrosoftTeams.d.ts` pour Typescript intellisense.

4. Nous avons ajouté un `manifest.dev.json`, `manifest.prod.json` et deux logos de taille *44x44* et *88x88*. N’oubliez pas de les renommer `manifest.json` dans vos fichiers zip que vous téléchargez dans Microsoft Teams.

5. Nous avons ajouté quelques styles à notre`app. css`.

6. Pour finir, nous avons modifié l’`authentification` dans outlook.tasks.ts pour qu'elle dépende plutôt de « useMicrosoftTeamsAuth », une nouvelle fonctionnalité de la version bêta de OfficeHelpers référencée dans cet exemple.

### Appel de la boîte de dialogue d’authentification

Lorsque l’utilisateur ajoute l’onglet, la page Configuration est présentée (config.html). Dans ce cas, le code authentifie l’utilisateur si possible.

L’authentification utilise une nouvelle fonction propre aux équipes dans la dernière version (>0.4.0) de la bibliothèque [Office-js-helpers](https://github.com/OfficeDev/office-js-helpers). Cette fonction d’assistance tentera d’effectuer l’authentification silencieuse, mais si ce n’est pas le cas, elle appelle la boîte de dialogue authentification Microsoft Teams proprement dit. Pour plus d’informations sur le processus d’authentification complet dans Microsoft Teams, consultez [Authentification d’un utilisateur](https://msdn.microsoft.com/en-us/microsoft-teams/auth) dans la [documentation des développeurs](https://msdn.microsoft.com/en-us/microsoft-teams/index) Microsoft Teams.

Dans outlook.tasks.ts :
```typescript
 return this.authenticator.authenticate('Microsoft')
    .then(token => this._token = token)
    .catch(error => {
        Utilities.log(error);
        throw new Error('Failed to login using your Microsoft Account');
    });
```

### Gestion de l’événement « Enregistrer »

Une fois la connexion établie, l’utilisateur enregistre l’onglet dans le canal. Le code suivant active le bouton Enregistrer et définit la fonction SaveHandler, qui stocke le contenu à afficher dans l’onglet (dans le cas présent, il s’agit uniquement du fichier index.html du projet).

Dans config.tsx :
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

## Technologie utilisée

Cet exemple utilises la pile suivant :

1. [`Réagir par Facebook`](https://facebook.github.io/react/) en tant qu’infrastructure d’interface utilisateur.
2. [`TypeScript`](https://www.typescriptlang.org/) comme transpileur.
4. [`TodoMVC`](http://todomvc.com/examples/typescript-react/#/) base de la fonctionnalité TodoMVC.
5. [`Webpack`](https://webpack.github.io/) comme outil de génération.

## Ressources supplémentaires

### Inscription d’une application pour l’authentification auprès de Microsoft
Notez que cela ne sera pas nécessaire si vous utilisez l'option Dev local ci-dessus, mais si vous choisissez d'héberger cet onglet dans votre propre environnement, vous devez inscrire l'application pour vous authentifier.

1. Accédez au [Portail d’inscription des applications de Microsoft](https://apps.dev.microsoft.com).
2. Connectez-vous à l’aide de votre compte Office 365 professionnel ou scolaire. N’utilisez pas votre compte Microsoft personnel.
2. Ajouter d’une nouvelle application.
2. Prenez note de votre nouveau `ID d’application`.
2. Cliquez sur `Ajouter une plateforme` et choisissez `Web`.
3. Cochez la case `Autoriser de flux implicite` et configurez l’URL de redirection pour qu’elle soit `https://<mywebsite>/config.html`.

Pour plus d’informations sur l’hébergement de pages d’onglets, consultez le fichier [Microsoft Teams « Prise en main » exemple de fichier LISEZ-MOI](https://github.com/OfficeDev/microsoft-teams-sample-get-started#host-tab-pages-over-https).

>**Remarque :** Par défaut, votre organisation doit vous autoriser à créer des applications. En revanche, si ce n’est pas le cas, vous pouvez demander à un développeur de tester gratuitement un abonnement d’un an à Office 365 Developer. [Voici comment procéder](https://msdn.microsoft.com/en-us/microsoft-teams/setup).


## Crédits

Ce projet est basé sur le modèle TodoMVC Typescript - React situé [ici](https://github.com/tastejs/todomvc/tree/gh-pages/examples/typescript-react).

## Contribution

Consultez [Contribution](contributing.md) pour en savoir plus sur notre code de conduite et sur le processus de soumission des demandes de réception.

Ce projet a adopté le [code de conduite Open Source de Microsoft](https://opensource.microsoft.com/codeofconduct/). Pour en savoir plus, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.

## Gestion des versions

Nous utilisons [SemVer](http://semver.org/) pour le contrôle des versions. Pour consulter les versions disponibles, consultez les [étiquettes dans ce référentiel](https://github.com/officedev/microsoft-teams-sample-todo/tags).

## Licence

Ce projet a fait l’objet d’une licence en vertu de la licence MIT-consultez le fichier de [Licence](LICENSE) pour plus de détails
