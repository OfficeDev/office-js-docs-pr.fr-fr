# <a name="build-your-first-word-add-in"></a>Cr?er votre premier compl?ment Word

_S?applique ? : Word 2016, Word pour iPad, Word pour Mac_

Cet article d?crit le processus de cr?ation d?un compl?ment Word ? l?aide de jQuery et de l?API JavaScript pour Word. 

## <a name="create-the-add-in"></a>Cr?er le compl?ment 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[Visual Studio](#tab/visual-studio)

### <a name="prerequisites"></a>Conditions pr?alables

[!include[Quickstart prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a>Cr?ation du projet de compl?ment

1. Dans la barre de menu de Visual Studio, choisissez successivement **Fichier** > **Nouveau** > **Projet**.
    
2. Dans la liste des types de projets, sous **Visual C#** ou **Visual Basic**, d?veloppez **Office/SharePoint**, choisissez **Compl?ments**, puis **Compl?ment Word Web** pour le type de projet. 

3. Nommez le projet, puis cliquez sur **OK**.

4. Visual Studio cr?e une solution et ses deux projets apparaissent dans l?**explorateur de solutions**. Le fichier **Home.html** s?ouvre dans Visual Studio.
    
### <a name="explore-the-visual-studio-solution"></a>Explorer la solution Visual Studio

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a>Mise ? jour du code

1. **Home.html** sp?cifie le code HTML qui s?affichera dans le volet Office du compl?ment. Dans **Home.html**, remplacez l??l?ment `<body>` par le balisage suivant et enregistrez le fichier.
 
    ```html
    <body>
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>    
        <div id="content-main">
            <div class="padding">
                <p>Choose the buttons below to add boilerplate text to the document by using the Word JavaScript API.</p>
                <br />
                <h3>Try it out</h3>
                <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                <br /><br />
                <button id="checkhov">Add quote from Anton Chekhov</button>
                <br /><br />
                <button id="proverb">Add Chinese proverb</button>
            </div>
        </div>
        <br />
        <div id="supportedVersion"/>
    </body>
    ```

2. Ouvrez le fichier **Home.js** ? la racine du projet d?application web. Ce fichier sp?cifie le script pour le compl?ment. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.

    ```js
    'use strict';
    
    (function () {

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

3. Ouvrez le fichier **Home.css** ? la racine du projet d?application web. Ce fichier sp?cifie les styles personnalis?s pour le compl?ment. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

### <a name="update-the-manifest"></a>Mise ? jour du manifeste

1. Ouvrez le fichier manifeste XML dans le projet de compl?ment. Ce fichier d?finit les param?tres et les fonctionnalit?s du compl?ment.

2. L??l?ment `ProviderName` poss?de une valeur d?espace r?serv?. Remplacez-le par votre nom.

3. L?attribut `DefaultValue` de l??l?ment `DisplayName` poss?de un espace r?serv?. Remplacez-le par **My Office Add-in**.

4. L?attribut `DefaultValue` de l??l?ment `Description` poss?de un espace r?serv?. Remplacez-le par **A task pane add-in for Word**.

5. Enregistrez le fichier.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="try-it-out"></a>Essayez !

1. ? l?aide de Visual Studio, testez le nouveau compl?ment en appuyant sur F5 ou en choisissant le bouton **D?marrer** pour lancer Word avec le bouton du compl?ment **Show Taskpane** (Afficher le volet Office) qui appara?t dans le ruban. Le compl?ment sera h?berg? localement sur IIS.

2. Dans Word, s?lectionnez l?onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du compl?ment.

    ![Capture d??cran de l?application Word avec le bouton Afficher le volet Office mis en ?vidence](../images/word-quickstart-addin-0.png)

3. Dans le volet Office, choisissez l?un des boutons pour ajouter du texte r?utilisable dans le document.

    ![Capture d??cran de l?application Word avec le compl?ment de texte r?utilisable charg?.](../images/word-quickstart-addin-1b.png)

# <a name="any-editortabvisual-studio-code"></a>[Tous les ?diteurs](#tab/visual-studio-code)

### <a name="prerequisites"></a>Conditions pr?alables

- [Node.js](https://nodejs.org)

- Installez la derni?re version de [Yeoman](https://github.com/yeoman/yo) et le [g?n?rateur Yeoman pour les compl?ments Office](https://github.com/OfficeDev/generator-office) globalement.

    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-add-in-project"></a>Cr?ation du projet de compl?ment

1. Cr?ez un dossier sur votre lecteur local et nommez-le `my-word-addin`. Il s?agit de l?emplacement dans lequel vous allez cr?er les fichiers de votre compl?ment.

2. Acc?dez ? votre nouveau dossier.

    ```bash
    cd my-word-addin
    ```

3. Utilisez le g?n?rateur Yeoman afin de cr?er un projet de compl?ment Word. Ex?cutez la commande suivante, puis r?pondez aux invites comme suit :

    ```bash
    yo office
    ```

    - **Voulez-vous cr?er un sous-dossier de votre projet ? :** `No`
    - **Comment souhaitez-vous nommer votre compl?ment ? :** `My Office Add-in`
    - **Quelle application client Office voulez-vous prendre en charge ? :** `Word`
    - **Voulez-vous cr?er un compl?ment ? :** `Yes`
    - **Souhaitez-vous utiliser TypeScript ? :** `No`
    - **Choisissez une infrastructure :** `Jquery`

    Le g?n?rateur demande ensuite si vous voulez ouvrir **resource.html**. Il n?est pas n?cessaire de l?ouvrir pour ce didacticiel, mais n?h?sitez pas ? l?ouvrir si vous ?tes curieux. Cliquez sur Oui ou Non pour fermer l?assistant et laisser le g?n?rateur faire son travail.

    ![Capture d??cran des invites et des r?ponses relatives au g?n?rateur Yeoman](../images/yo-office-word-jquery.png)

### <a name="update-the-code"></a>Mise ? jour du code

1. Dans votre ?diteur de code, ouvrez **index.html** ? la racine du projet. Ce fichier contient le code HTML qui s?affichera dans le volet Office du compl?ment. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier. Ce compl?ment affichera trois boutons et, lorsque l?un d?eux sera choisi, du texte r?utilisable sera ajout? au document.

    ```html
    <!DOCTYPE html>
    <html>
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
            <title>Boilerplate text app</title>
            <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>
            <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
            <script src="app.js" type="text/javascript"></script>
            <link href="app.css" rel="stylesheet" type="text/css" />
        </head>
        <body>
            <div id="content-header">
                <div class="padding">
                    <h1>Welcome</h1>
                </div>
            </div>    
            <div id="content-main">
                <div class="padding">
                    <p>Choose the buttons below to add boilerplate text to the document by using the Word JavaScript API.</p>
                    <br />
                    <h3>Try it out</h3>
                    <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                    <br /><br />
                    <button id="checkhov">Add quote from Anton Chekhov</button>
                    <br /><br />
                    <button id="proverb">Add Chinese proverb</button>
                </div>
            </div>
            <br />
            <div id="supportedVersion"/>
        </body>
    </html>
    ```

2. Ouvrez le fichier **app.js** pour sp?cifier le script pour le compl?ment. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier. Ce script contient le code d?initialisation ainsi que le code qui apporte des modifications au document Word en ins?rant du texte dans le document lorsqu?un bouton est choisi. 

    ```js
    'use strict';
    
    (function () {

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

3. Ouvrez le fichier **app.css** ? la racine du projet pour sp?cifier les styles personnalis?s du compl?ment. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

### <a name="update-the-manifest"></a>Mise ? jour du manifeste

1. Ouvrez le fichier nomm? **my-office-add-in-manifest.xml** pour d?finir les param?tres et les fonctionnalit?s du compl?ment.

2. L??l?ment `ProviderName` poss?de une valeur d?espace r?serv?. Remplacez-le par votre nom.

3. L?attribut `DefaultValue` de l??l?ment `Description` poss?de un espace r?serv?. Remplacez-le par **A task pane add-in for Word**.

4. Enregistrez le fichier.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="start-the-dev-server"></a>D?marrage du serveur de d?veloppement

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a>Essayez !

1. Suivez les instructions pour la plateforme que vous utiliserez afin d?ex?cuter votre compl?ment en vue d?en charger une version test dans Word.

    - Windows : [Chargement de version test des compl?ments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Word Online : [Chargement d?une version test des compl?ments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad et Mac : [Chargement de version test des compl?ments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

2. Dans Word, s?lectionnez l?onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du compl?ment.

    ![Capture d??cran de l?application Word avec le bouton Afficher le volet Office mis en ?vidence](../images/word-quickstart-addin-2.png)

3. Dans le volet Office, choisissez l?un des boutons pour ajouter du texte r?utilisable dans le document.

    ![Capture d??cran de l?application Word avec le compl?ment de texte r?utilisable charg?.](../images/word-quickstart-addin-1.png)

---

## <a name="next-steps"></a>?tapes suivantes

F?licitations, vous avez cr?? un compl?ment Word ? l?aide de jQuery ! Ensuite, d?couvrez les fonctionnalit?s d?un compl?ment Word et cr?ez-en un plus complexe en suivant le didacticiel sur les compl?ments Word.

> [!div class="nextstepaction"]
> [Didacticiel sur les compl?ments Word](../tutorials/word-tutorial.yml)

## <a name="see-also"></a>Voir aussi

* [Pr?sentation des compl?ments Word](../word/word-add-ins-programming-overview.md)
* [Exemples de code pour les compl?ments Word](http://dev.office.com/code-samples#?filters=word,office%20add-ins)
* [R?f?rence d?API JavaScript pour Word](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview)
