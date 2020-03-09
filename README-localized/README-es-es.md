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
# Ejemplo de aplicación de pestaña “Lista de tareas pendientes” de Microsoft Teams
[![Estado de la compilación](https://travis-ci.org/OfficeDev/microsoft-teams-sample-todo.svg?branch=master)](https://travis-ci.org/OfficeDev/microsoft-teams-sample-todo)

Esta es una [aplicación de pestaña de ejemplo para Microsoft Teams](https://aka.ms/microsoftteamstabsplatform). La finalidad de este ejemplo es mostrar lo fácil que es convertir una aplicación web existente en una aplicación de pestaña de Microsoft Teams. La aplicación web existente, [**TodoMVC para React**](https://github.com/tastejs/todomvc/tree/gh-pages/examples/typescript-react), ofrece un administrador de tareas básico que se integra con las tareas personales de Outlook. Con solo unos pocos cambios menores, esta vista web se puede agregar a un canal como una aplicación de pestañas. Eche un vistazo al [código diff entre las ramas "Before" y "after"](https://github.com/OfficeDev/microsoft-teams-sample-todo/compare/before...after) para ver los cambios que se hicieron.

> **Nota:** Este es un ejemplo realista de una aplicación de colaboración en equipo. Las tareas que se muestran pertenecen a la cuenta individual del usuario y no a una cuenta de equipo compartida.

**Para obtener más información sobre cómo desarrollar experiencias para Microsoft Teams, consulte la [documentación para desarrolladores](https://msdn.microsoft.com/en-us/microsoft-teams/index) de Microsoft Teams.**

## Requisitos previos

1. Un [cuenta de Office 365 con acceso a Microsoft Teams](https://msdn.microsoft.com/en-us/microsoft-teams/setup).
2. Este ejemplo se ha creado con [Node.js](https://nodejs.org). Descargue e instale la versión recomendada si todavía no la tiene.

## Ejecute la aplicación

### Hospede las aplicaciones de pestaña "página de configuración" y "página de contenido"

Ya que la configuración de pestaña y las páginas de contenido son aplicaciones web, deben estar hospedadas.  

En primer lugar, prepararemos el proyecto:

1. Clone el repositorio.
2. Abra una línea de comandos en el subdirectorio de repositorios.
3. Ejecute `npm install` (incluido como parte de node. js) desde la línea de comandos.

Para hospedar la aplicación de forma local:
* en el subdirectorio de repositorio, ejecute `npm start` para iniciar el `webpack-dev-server` para habilitar la aplicación dev para que funcione.
* Tenga en cuenta que el manifiesto dev está configurado para usar localhost, por lo que no se necesitaría ninguna configuración adicional.

De forma opcional, para crear una aplicación que se puede implementar y hospedar en su propio entorno:
* en el subdirectorio de repositorios, ejecute `npm run build` para generar una compilación implementable.
* Modifique config.url en el manifiesto del producto para que apunte a la ubicación de hospedaje. [Vea lo siguiente](#registering-an-application-to-authenticate-with-microsoft)

### Agregue una pestaña a Microsoft Teams

1. Descargue el archivo comprimido de [paquete de desarrollo de aplicaciones de pestaña](https://github.com/OfficeDev/microsoft-teams-sample-todo/tree/master/package/todo.dev.zip) para este ejemplo.
2. Cree un nuevo equipo para probarlo si es necesario. Haga clic en **crear equipo** en la parte inferior del panel izquierdo.
3. Seleccione el equipo en el panel izquierdo, seleccione **... (más opciones)** y seleccione **ver equipo**.
4. Seleccione la pestaña **desarrollador (vista previa)** y seleccione **cargar**.
5. Desplácese hasta el archivo comprimido que ha descargado en el paso 1 anterior y selecciónelo.
6. Vaya a cualquier canal del equipo. Haga clic en el «+» a la derecha de las pestañas existentes.
7. Seleccione la pestaña de la galería que aparece.
8. Acepte el consentimiento.
9. Si es necesario, inicie sesión con su cuenta profesional o educativa de Office 365. Tenga en cuenta que, si es posible, el código intentará realizar la autenticación silenciosa.
10. Valide la información de autenticación.
11. Presione guardar para agregar la pestaña al canal.

> **Nota:** Para volver a cargar un paquete actualizado con el mismo `id`, haga clic en el icono 'reemplazar ' al final de la fila de tabla de la pestaña. No vuelva a hacer clic en "cargar": Microsoft Teams dirá que la pestaña ya existe.

> Es recomendable tener varias configuraciones, una por entorno. Los nombres de los archivos zip pueden ser cualquiera, como, por ejemplo,`todo.dev.zip`, `todo.prod.zip`, etc., pero el archivo comprimido debe contener un `manifest.json` con un `id` único. Consulte [crear un manifiesto para su pestaña](https://msdn.microsoft.com/en-us/microsoft-teams/createpackage).

## Recorrido del código

Mientras que la rama `principal` muestra el estado más reciente de la muestra, eche un vistazo a las siguientes [diferencias de código](https://github.com/OfficeDev/microsoft-teams-sample-todo/compare/before...after) entre:

* [`before`](https://github.com/OfficeDev/microsoft-teams-sample-todo/commit/before): la aplicación inicial

* [`after`](https://github.com/OfficeDev/microsoft-teams-sample-todo/commit/after): la aplicación después de la integración con Microsoft Teams.

Siguiendo esto paso a paso:

1. Se ha agregado una nueva página de `config.html` y `config.tsx` que es responsable de la aplicación para permitir que el usuario manipule la configuración, realice la autenticación de inicio de sesión único, etc. durante el primer inicio. Esto es necesario para que el administrador de equipo pueda establecer la aplicación o la configuración. Consulte [crear la página de configuración](https://msdn.microsoft.com/en-us/microsoft-teams/createconfigpage).

2. Hemos agregado el mismo archivo `config.html` a nuestra configuración `webpack.common.js` para que pueda inyectar los paquetes adecuados durante el tiempo de ejecución.

3. Hemos agregado una referencia a `MicrosoftTeams.js` en index.html y `MicrosoftTeams.d.ts` para Typescript IntelliSense.

4. Hemos agregado `manifest.dev.json`, `manifest.prod.json` y dos logotipos para el tamaño *44x44* y *88x88*. Recuerde cambiar el nombre de `manifest.json` en los archivos zip que cargue en Microsoft Teams.

5. Hemos agregado ciertos estilos a nuestra `app.css`.

6. Por último, hemos modificado la `autenticación` en outlook.tasks.ts para depender en lugar de "useMicrosoftTeamsAuth", una nueva característica de la versión beta de OfficeHelpers a la que se hace referencia en este ejemplo.

### Invocar el cuadro de diálogo autenticación

Cuando el usuario agrega la pestaña, se muestra la página de configuración (config.html). En este caso, el código autenticará al usuario, si es posible.

La autenticación usa una función específica para Teams en la versión más reciente (> 0.4.0) de la biblioteca [office-js-helpers](https://github.com/OfficeDev/office-js-helpers). Esta función auxiliar intentará autenticar de forma silenciosa, pero si no lo hace, llamará al diálogo de autenticación específica de Microsoft Teams. Para obtener más información sobre el proceso de autenticación completo en Microsoft Teams, consulte [autenticación de un usuario](https://msdn.microsoft.com/en-us/microsoft-teams/auth) en la [documentación para desarrolladores](https://msdn.microsoft.com/en-us/microsoft-teams/index) de Microsoft Teams.

En outlook.tasks.ts:
```typescript
 return this.authenticator.authenticate('Microsoft')
    .then(token => this._token = token)
    .catch(error => {
        Utilities.log(error);
        throw new Error('Failed to login using your Microsoft Account');
    });
```

### Controlar el evento "guardar"

Una vez que se haya iniciado sesión correctamente, el usuario guardará la pestaña en el canal. El código siguiente habilita el botón Guardar, y establece SaveHandler, que almacenará el contenido que se muestra en la pestaña (en este caso, solo el index.html del proyecto).

En config.tsx:
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

## Tecnología usada

Este ejemplo usa:

1. [`React por Facebook`](https://facebook.github.io/react/) como marco de trabajo de la interfaz de usuario.
2. [`TypeScript`](https://www.typescriptlang.org/) como transpilador.
4. [`TodoMVC`](http://todomvc.com/examples/typescript-react/#/) base para la funcionalidad de TodoMVC.
5. [`WebPack`](https://webpack.github.io/) como herramienta de compilación.

## Recursos adicionales

### Registrar una aplicación para autenticar con Microsoft
Tenga en cuenta que esto no será necesario si usa la opción de desarrollo local anterior, pero si decide hospedar esta pestaña en su propio entorno, debe registrar la aplicación para autenticar.

1. Vaya al [Portal de registro de aplicaciones de Microsoft](https://apps.dev.microsoft.com).
2. Inicie sesión con su cuenta profesional o educativa de Office 365. No use una cuenta de Microsoft personal.
2. Agregue una nueva aplicación.
2. Tome nota del nuevo `identificador de la aplicación`.
2. Haga clic en `agregar plataforma` y elija `web`.
3. Marque `permitir flujo implícito` y configure la dirección URL de redireccionamiento para que sea `https://<mywebsite>/config.html`.

Para obtener más información sobre cómo hospedar sus propias páginas de pestañas, consulte el [archivo Léame de ejemplo de introducción a Microsoft Teams](https://github.com/OfficeDev/microsoft-teams-sample-get-started#host-tab-pages-over-https).

>**Nota:** De forma predeterminada, la organización debería permitir crear aplicaciones nuevas. Sin embargo, si no lo hace, puede obtener un espacio empresarial de prueba para desarrolladores, una suscripción de prueba de un año de Office 365 Developer gratis. [Le mostramos cómo](https://msdn.microsoft.com/en-us/microsoft-teams/setup).


## Créditos

Este proyecto se basa en la plantilla TodoMVC Typescript - React ubicada [aquí](https://github.com/tastejs/todomvc/tree/gh-pages/examples/typescript-react).

## Colaboradores

Lea [colaboradores](contributing.md) para obtener más información sobre nuestro código de conducta y sobre el proceso de envío de solicitudes de incorporación de cambios.

Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, vea [Preguntas frecuentes sobre el código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.

## Control de versiones

Usamos [SemVer](http://semver.org/) para el control de versiones. Para ver las versiones disponibles, consulte las [etiquetas de este repositorio](https://github.com/officedev/microsoft-teams-sample-todo/tags).

## Licencia

Este proyecto está publicado bajo la licencia MIT. Consulte el archivo de [licencia](LICENSE) para obtener más información.
