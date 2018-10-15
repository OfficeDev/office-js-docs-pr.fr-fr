# <a name="build-your-first-word-add-in"></a><span data-ttu-id="eaf1b-101">Créer votre premier complément Word</span><span class="sxs-lookup"><span data-stu-id="eaf1b-101">Build your first Word add-in</span></span>

<span data-ttu-id="eaf1b-102">_S’applique à : Word 2016, Word pour iPad, Word pour Mac_</span><span class="sxs-lookup"><span data-stu-id="eaf1b-102">_Applies to: Word 2016, Word for iPad, Word for Mac_</span></span>

<span data-ttu-id="eaf1b-103">Cet article décrit le processus de création d’un complément Word à l’aide de jQuery et de l’API JavaScript pour Word.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-103">In this article, you'll walk through the process of building a Word add-in by using jQuery and the Word JavaScript API.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="eaf1b-104">Créer le complément</span><span class="sxs-lookup"><span data-stu-id="eaf1b-104">Create the add-in</span></span> 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[<span data-ttu-id="eaf1b-105">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="eaf1b-105">Visual Studio</span></span>](#tab/visual-studio)

### <a name="prerequisites"></a><span data-ttu-id="eaf1b-106">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="eaf1b-106">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="eaf1b-107">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="eaf1b-107">Create the add-in project</span></span>

1. <span data-ttu-id="eaf1b-108">Dans la barre de menu de Visual Studio, choisissez successivement **Fichier** > **Nouveau** > **Projet**.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-108">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="eaf1b-109">Dans la liste des types de projets, sous **Visual C#** ou **Visual Basic**, développez **Office/SharePoint**, choisissez **Compléments**, puis **Complément Word Web** pour le type de projet.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-109">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Word Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="eaf1b-110">Nommez le projet, puis cliquez sur **OK**.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-110">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="eaf1b-p101">Visual Studio crée une solution et ses deux projets apparaissent dans l’**Explorateur de solutions**. Le fichier **Home.html** s’ouvre dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>
    
### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="eaf1b-113">Explorer la solution Visual Studio</span><span class="sxs-lookup"><span data-stu-id="eaf1b-113">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="eaf1b-114">Mise à jour du code</span><span class="sxs-lookup"><span data-stu-id="eaf1b-114">Update the code</span></span>

1. <span data-ttu-id="eaf1b-p102">**Home.html** spécifie le code HTML qui s’affichera dans le volet Office du complément. Dans **Home.html**, remplacez l’élément `<body>` par le balisage suivant et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-p102">**Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
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

2. <span data-ttu-id="eaf1b-p103">Ouvrez le fichier **Home.js** à la racine du projet d’application web. Ce fichier spécifie le script pour le complément. Remplacez l'intégralité du contenu par le code suivant et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-p103">Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.</span></span>

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
                    $('#supportedVersion').html('This code is using Word 2016 or later.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or later.');
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

3. <span data-ttu-id="eaf1b-p104">Ouvrez le fichier **Home.css** à la racine du projet d’application web. Ce fichier spécifie les styles personnalisés pour le complément. Remplacez l'intégralité du contenu par le code suivant et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-p104">Open the file **Home.css** in the root of the web application project. This file specifies the custom styles for the add-in. Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="eaf1b-123">Mise à jour du manifeste</span><span class="sxs-lookup"><span data-stu-id="eaf1b-123">Update the manifest</span></span>

1. <span data-ttu-id="eaf1b-p105">Ouvrez le fichier manifeste XML dans le projet de complément. Ce fichier définit les paramètres et les fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-p105">Open the XML manifest file in the add-in project. This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="eaf1b-p106">L’élément `ProviderName` possède une valeur d’espace réservé. Remplacez-la par votre nom.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-p106">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="eaf1b-p107">L’attribut `DefaultValue` de l’élément `DisplayName` possède un espace réservé. Remplacez-le par **Mon complément Office**.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-p107">The `DefaultValue` attribute of the `DisplayName` element has a placeholder. Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="eaf1b-p108">L’attribut `DefaultValue`  de l’élément `Description`  possède un espace réservé. Remplacez-le par **un complément volet Office pour Word**.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-p108">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for Word**.</span></span>

5. <span data-ttu-id="eaf1b-132">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-132">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="eaf1b-133">Essayez-le</span><span class="sxs-lookup"><span data-stu-id="eaf1b-133">Try it out</span></span>

1. <span data-ttu-id="eaf1b-p109">À l’aide de Visual Studio, testez le nouveau complément en appuyant sur F5 ou en choisissant le bouton **Démarrer** pour lancer Word avec le bouton du complément **Afficher le volet Office** qui apparaît dans le ruban. Le complément sera hébergé localement sur IIS.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-p109">Using Visual Studio, test the newly created Word add-in by pressing F5 or choosing the **Start** button to launch Word with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="eaf1b-136">Dans Word, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-136">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="eaf1b-137">(Si vous utilisez la version sans abonnement d’Office 2016, au lieu de la version d’Office 365, les boutons personnalisés ne sont pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-137">(If you are using the non-subscription version of Office 2016, instead of the Office 365 version, then custom buttons are not supported.</span></span> <span data-ttu-id="eaf1b-138">Au lieu de cela, le volet Office s’ouvre immédiatement.)</span><span class="sxs-lookup"><span data-stu-id="eaf1b-138">Instead, the task pane will open immediately.)</span></span>

    ![Capture d’écran de l’application Word avec le bouton Afficher le volet Office mis en surbrillance](../images/word-quickstart-addin-0.png)

3. <span data-ttu-id="eaf1b-140">Dans le volet Office, choisissez l’un des boutons pour ajouter du texte réutilisable dans le document.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-140">In the task pane, choose any of the buttons to add boilerplate text to the document.</span></span>

    ![Capture d’écran de l’application Word avec le complément de texte réutilisable chargé.](../images/word-quickstart-addin-1b.png)

# <a name="any-editortabvisual-studio-code"></a>[<span data-ttu-id="eaf1b-142">Tous les éditeurs</span><span class="sxs-lookup"><span data-stu-id="eaf1b-142">Any editor</span></span>](#tab/visual-studio-code)

### <a name="prerequisites"></a><span data-ttu-id="eaf1b-143">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="eaf1b-143">Prerequisites</span></span>

- [<span data-ttu-id="eaf1b-144">Node.js</span><span class="sxs-lookup"><span data-stu-id="eaf1b-144">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="eaf1b-145">Installez la dernière version de [Yeoman](https://github.com/yeoman/yo) et le [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) globalement.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-145">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-add-in-project"></a><span data-ttu-id="eaf1b-146">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="eaf1b-146">Create the add-in project</span></span>

1. <span data-ttu-id="eaf1b-147">Créez un dossier sur votre lecteur local et nommez-le `my-word-addin`.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-147">Create a folder on your local drive and name it `my-word-addin`.</span></span> <span data-ttu-id="eaf1b-148">Il s’agit de l’emplacement dans lequel vous allez créer les fichiers de votre complément.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-148">This is where you'll create the files for your add-in.</span></span>

2. <span data-ttu-id="eaf1b-149">Accédez à votre nouveau dossier.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-149">Navigate to your new folder.</span></span>

    ```bash
    cd my-word-addin
    ```

3. <span data-ttu-id="eaf1b-150">Utilisez le générateur Yeoman afin de créer un projet de complément Word.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-150">Use the Yeoman generator to create a Word add-in project.</span></span> <span data-ttu-id="eaf1b-151">Exécutez la commande suivante, puis répondez aux invites comme suit :</span><span class="sxs-lookup"><span data-stu-id="eaf1b-151">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="eaf1b-152">**Choisissez un type de projet :** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="eaf1b-152">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="eaf1b-153">**Choisissez un type de script :** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="eaf1b-153">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="eaf1b-154">**Comment souhaitez-vous nommer votre complément ?** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="eaf1b-154">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="eaf1b-155">**Quelle application client Office voulez-vous prendre en charge ?** `Word`</span><span class="sxs-lookup"><span data-stu-id="eaf1b-155">**Which Office client application would you like to support?:** `Word`</span></span>

    ![Capture d’écran des invites et des réponses pour le générateur Yeoman](../images/yo-office-word-jquery.png)
    
    <span data-ttu-id="eaf1b-157">Une fois que vous avez terminé avec l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-157">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>
    
4. <span data-ttu-id="eaf1b-158">Accédez au dossier racine du projet d’application web.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-158">Navigate to the root folder of the web application project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

### <a name="update-the-code"></a><span data-ttu-id="eaf1b-159">Mise à jour du code</span><span class="sxs-lookup"><span data-stu-id="eaf1b-159">Update the code</span></span>

1. <span data-ttu-id="eaf1b-160">Dans votre éditeur de code, ouvrez **index.html** à la racine du projet.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-160">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="eaf1b-161">Ce fichier contient le code HTML qui s’affichera dans le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-161">This file contains the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="eaf1b-162">Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-162">Replace the entire contents with the following code and save the file.</span></span> <span data-ttu-id="eaf1b-163">Ce complément affichera trois boutons et, lorsque l’un d’eux sera choisi, du texte réutilisable sera ajouté au document.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-163">This add-in will display three buttons and when any of the buttons are chosen, boilerplate text will be added to the document.</span></span>

    ```html
    <!DOCTYPE html>
    <html>
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
            <title>Boilerplate text app</title>
            <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>
            <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
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

2. <span data-ttu-id="eaf1b-164">Ouvrez le fichier **src/index.js** pour spécifier le script du complément.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-164">Open the file **app.js** to specify the script for the add-in.</span></span> <span data-ttu-id="eaf1b-165">Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-165">Replace the entire contents with the following code and save the file.</span></span> <span data-ttu-id="eaf1b-166">Ce script contient le code d’initialisation ainsi que le code qui apporte des modifications au document Word en insérant du texte dans le document lorsqu’un bouton est choisi.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-166">This script contains initialization code as well as the code that makes changes to the Word document, by inserting text into the document when a button is chosen.</span></span> 

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
                    $('#supportedVersion').html('This code is using Word 2016 or later.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or later.');
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

3. <span data-ttu-id="eaf1b-167">Ouvrez le fichier **app.css** à la racine du projet pour spécifier les styles personnalisés du complément.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-167">Open the file **app.css** in the root of the project to specify the custom styles for the add-in.</span></span> <span data-ttu-id="eaf1b-168">Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-168">Replace the entire contents with the following and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="eaf1b-169">Mise à jour du manifeste</span><span class="sxs-lookup"><span data-stu-id="eaf1b-169">Update the manifest</span></span>

1. <span data-ttu-id="eaf1b-170">Ouvrez le fichier nommé **my-office-add-in-manifest.xml** pour définir les paramètres et les fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-170">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="eaf1b-171">L’élément `ProviderName` possède une valeur d’espace réservé.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-171">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="eaf1b-172">Remplacez-le par votre nom.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-172">Replace it with your name.</span></span>

3. <span data-ttu-id="eaf1b-173">L’attribut `DefaultValue` de l’élément `Description` possède un espace réservé.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-173">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="eaf1b-174">Remplacez-le par **Un complément de volet Office pour Word**.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-174">Replace it with **A task pane add-in for Word**.</span></span>

4. <span data-ttu-id="eaf1b-175">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-175">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="start-the-dev-server"></a><span data-ttu-id="eaf1b-176">Démarrage du serveur de développement</span><span class="sxs-lookup"><span data-stu-id="eaf1b-176">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a><span data-ttu-id="eaf1b-177">Essayez-le</span><span class="sxs-lookup"><span data-stu-id="eaf1b-177">Try it out</span></span>

1. <span data-ttu-id="eaf1b-178">Suivez les instructions pour la plateforme que vous utiliserez afin d’exécuter votre complément en vue d’en charger une version test dans Word.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-178">To sideload the add-in within Word, follow the instructions for the platform you'll use to run your add-in.</span></span>

    - <span data-ttu-id="eaf1b-179">Windows : [Chargement d’une version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="eaf1b-179">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="eaf1b-180">Word Online : [Chargement d’une version test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="eaf1b-180">Word Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="eaf1b-181">iPad et Mac : [Chargement d’une version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="eaf1b-181">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="eaf1b-182">Dans Word, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-182">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Capture d’écran de l’application Word avec le bouton Afficher le volet Office mis en surbrillance](../images/word-quickstart-addin-2.png)

3. <span data-ttu-id="eaf1b-184">Dans le volet Office, choisissez l’un des boutons pour ajouter du texte réutilisable dans le document.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-184">In the task pane, choose any of the buttons to add boilerplate text to the document.</span></span>

    ![Capture d’écran de l’application Word avec le complément de texte réutilisable chargé.](../images/word-quickstart-addin-1.png)

---

## <a name="next-steps"></a><span data-ttu-id="eaf1b-186">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="eaf1b-186">Next steps</span></span>

<span data-ttu-id="eaf1b-187">Félicitations, vous avez créé un complément Word à l’aide de jQuery !</span><span class="sxs-lookup"><span data-stu-id="eaf1b-187">Congratulations, you've successfully created a Word add-in using jQuery!</span></span> <span data-ttu-id="eaf1b-188">Découvrez maintenant les fonctionnalités d’un complément Word et créez-en un plus complexe en suivant le didacticiel sur le complément Word.</span><span class="sxs-lookup"><span data-stu-id="eaf1b-188">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="eaf1b-189">Didacticiel sur les compléments Word</span><span class="sxs-lookup"><span data-stu-id="eaf1b-189">Word add-in tutorial</span></span>](../tutorials/word-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="eaf1b-190">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="eaf1b-190">See also</span></span>

* [<span data-ttu-id="eaf1b-191">Présentation des compléments Word</span><span class="sxs-lookup"><span data-stu-id="eaf1b-191">Word add-ins overview</span></span>](../word/word-add-ins-programming-overview.md)
* [<span data-ttu-id="eaf1b-192">Exemples de code pour les compléments Word</span><span class="sxs-lookup"><span data-stu-id="eaf1b-192">Word add-in code samples</span></span>](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,Word)
* [<span data-ttu-id="eaf1b-193">Référence d’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="eaf1b-193">Word JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/word-add-ins-reference-overview?view=office-js)
