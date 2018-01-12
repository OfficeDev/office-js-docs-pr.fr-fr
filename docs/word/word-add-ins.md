# <a name="build-your-first-word-add-in"></a>Créer votre premier complément Word

_S’applique à : Word 2016, Word pour iPad, Word pour Mac_

Un complément Word est exécuté à l’intérieur de Word et peut interagir avec le contenu du document à l’aide de l’API JavaScript pour Word, qui fait partie du modèle de programmation des compléments Office pour étendre des applications Office. Dans le modèle de programmation de ce complément, vous pouvez utiliser la plateforme et la langue de votre choix pour créer l’application web qui héberge votre extension à Word puis utiliser le [manifeste](../../docs/overview/add-in-manifests.md) du complément pour définir ses paramètres et fonctionnalités.

Cet article décrit le processus de création d’un complément Word à l’aide de jQuery et de l’API JavaScript pour Word. 

> **Remarque** : pour développer un complément pour Word 2013, vous devez utiliser l’[API Javascript pour Office]( https://dev.office.com/docs/add-ins/word/word-add-ins-programming-overview#javascript-apis-for-word) partagée. Pour en savoir plus sur les plateformes et les différentes API disponibles, reportez-vous à [Disponibilité des compléments Office sur les plateformes et les hôtes](https://dev.office.com/add-in-availability). 

## <a name="create-the-web-app"></a>Création de l’application web 

1. Créez un dossier sur votre lecteur local et nommez-le **BoilerplateAddin**. Il s’agit de l’endroit où vous allez créer les fichiers de votre application.

2. Dans le dossier de l’application, créez un fichier nommé **home.html** pour spécifier le code HTML qui sera affiché dans le volet Office du complément. Ce complément affichera trois boutons et, lorsque l’un d’eux sera choisi, du texte réutilisable sera ajouté au document. Ajoutez le code suivant à **home.html** et enregistrez le fichier.

    ```html
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Boilerplate text app</title>
        <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
        <script src="home.js" type="text/javascript"></script>
        </head>
        <body>
            <div>
                <h1>Welcome</h1>
            </div>
            <div>
                <p>This sample shows how to add boilerplate text to a document by using the Word JavaScript API.</p>
                <br />
                <h3>Try it out</h3>
                <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                <button id="checkhov">Add quote from Anton Chekhov</button>
                <button id="proverb">Add Chinese proverb</button>
            </div>
            <h3><div id="supportedVersion"/></h3>
        </body>
    </html>
    ```

3. Dans le dossier de l’application, créez un fichier nommé **home.js** pour spécifier le script jQuery du complément. Ce script contient le code d’initialisation ainsi que le code qui apporte des modifications au document Word en insérant du texte dans le document lorsqu’un bouton est choisi. Ajoutez le code suivant à **home.js** et enregistrez le fichier.

    ```javascript
    (function () {
        "use strict";

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {

                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#emerson').click(insertEmersonQuoteAtSelection);
                    $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                    $('#proverb').click(insertChineseProverbAtTheEnd);
                    $('#supportedVersion').html('This code is using Word 2016 or greater.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
                }
            });
        };

        function insertEmersonQuoteAtSelection() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();

                // Queue a command to replace the selected text.
                range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Ralph Waldo Emerson.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChekhovQuoteAtTheBeginning() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the start of the document body.
                body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Anton Chekhov.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChineseProverbAtTheEnd() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the end of the document body.
                body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from a Chinese proverb.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

## <a name="create-the-manifest-file"></a>Création du fichier manifeste

1. Dans le dossier de l’application, créez un fichier nommé **BoilerplateManifest.xml** pour définir les paramètres et les fonctionnalités du complément. Ajoutez le code suivant au fichier. 

    ```xml
    <?xml version="1.0" encoding="UTF-8"?>
        <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xsi:type="TaskPaneApp">
            <Id>2b88100c-656e-4bab-9f1e-f6731d86e464</Id>
            <Version>1.0.0.0</Version>
            <ProviderName>Microsoft</ProviderName>
            <DefaultLocale>en-US</DefaultLocale>
            <DisplayName DefaultValue="Boilerplate content" />
            <Description DefaultValue="Insert boilerplate content into a Word document." />
            <Hosts>
                <Host Name="Document"/>
            </Hosts>
            <DefaultSettings>
                <SourceLocation DefaultValue="\\MyShare\boilerplate\home.html" />
            </DefaultSettings>
            <Permissions>ReadWriteDocument</Permissions>
        </OfficeApp>
    ```

2. Générez un GUID à l’aide d’un générateur en ligne de votre choix. Ensuite, remplacez la valeur de l’élément **Id** indiquée à l’étape précédente par ce GUID.

3. Enregistrez le fichier manifeste.

## <a name="deploy-the-web-app-and-update-the-manifest"></a>Déployer l’application web et mettre à jour le manifeste

1. Déployez votre application web (par exemple, le contenu du dossier de votre application) sur le serveur web de votre choix.

2. Dans le dossier local de l’application, ouvrez le fichier manifeste (**BoilerplateManifest.xml**). Modifiez la valeur d’attribut dans l’élément **SourceLocation** pour spécifier l’emplacement du fichier **home.html** sur le serveur web et enregistrez le fichier.

## <a name="try-it-out"></a>Essayez !

1. Suivez les instructions pour la plateforme que vous utiliserez afin d’exécuter votre complément en vue d’en charger une version test dans Word.

    - Windows : [Chargement de compléments Office pour des tests sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Word Online : [Chargement d’une version test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

2. Dans le volet Office de droite, choisissez l’un des boutons pour ajouter du texte réutilisable dans le document.

![Image de l’application Word avec le complément boilerplate chargé.](../../images/boilerplateAddin.png)

## <a name="next-steps"></a>Étapes suivantes

Félicitations, vous avez créé un complément Word à l’aide de jQuery ! Apprenez-en davantage sur les [concepts fondamentaux](word-add-ins-programming-overview.md) de la création de compléments Word.

## <a name="additional-resources"></a>Ressources supplémentaires

* [Présentation des compléments Word](word-add-ins-programming-overview.md)
* [Exemples de code pour les compléments Word](http://dev.office.com/code-samples#?filters=word,office%20add-ins)
* [Référence d’API JavaScript pour Word](../../reference/word/word-add-ins-reference-overview.md)