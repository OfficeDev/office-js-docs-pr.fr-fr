---
title: Développement de votre premier complément Excel à l’aide de jQuery
description: ''
ms.date: 03/19/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: d1783d257b0d46c42f16cf5e1f261407a355a63d
ms.sourcegitcommit: 68872372d181cca5bee37ade73c2250c4a56bab6
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/01/2019
ms.locfileid: "33517078"
---
# <a name="build-an-excel-add-in-using-jquery"></a><span data-ttu-id="25bd5-102">Développement d’un complément Excel à l’aide de jQuery</span><span class="sxs-lookup"><span data-stu-id="25bd5-102">Build an Excel add-in using jQuery</span></span>

<span data-ttu-id="25bd5-103">Cet article décrit le processus de création d’un complément Excel à l’aide de jQuery et de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="25bd5-103">In this article, you'll walk through the process of building an Excel add-in by using jQuery and the Excel JavaScript API.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="25bd5-104">Créer le complément</span><span class="sxs-lookup"><span data-stu-id="25bd5-104">Create the add-in</span></span>

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[<span data-ttu-id="25bd5-105">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="25bd5-105">Visual Studio</span></span>](#tab/visual-studio)

### <a name="prerequisites"></a><span data-ttu-id="25bd5-106">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="25bd5-106">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="25bd5-107">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="25bd5-107">Create the add-in project</span></span>

1. <span data-ttu-id="25bd5-108">Dans la barre de menu de Visual Studio, choisissez successivement **Fichier** > **Nouveau** > **Projet**.</span><span class="sxs-lookup"><span data-stu-id="25bd5-108">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>

2. <span data-ttu-id="25bd5-109">Dans la liste des types de projet, sous **Visual C#** ou **Visual Basic**, développez **Office/SharePoint**, choisissez **Compléments**, puis **Complément Excel Web** pour le type de projet.</span><span class="sxs-lookup"><span data-stu-id="25bd5-109">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="25bd5-110">Nommez le projet, puis cliquez sur **OK**.</span><span class="sxs-lookup"><span data-stu-id="25bd5-110">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="25bd5-111">Dans la fenêtre de dialogue **Créer un complément Office**, sélectionnez **Ajouter de nouvelles fonctionnalités à Excel**, puis sélectionnez **Terminer** pour créer le projet.</span><span class="sxs-lookup"><span data-stu-id="25bd5-111">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="25bd5-p101">Visual Studio crée une solution et ses deux projets apparaissent dans l’**explorateur de solutions**. Le fichier **Home.html** s’ouvre dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="25bd5-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="25bd5-114">Explorer la solution Visual Studio</span><span class="sxs-lookup"><span data-stu-id="25bd5-114">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="25bd5-115">Mise à jour du code</span><span class="sxs-lookup"><span data-stu-id="25bd5-115">Update the code</span></span>

1. <span data-ttu-id="25bd5-p102">**Home.html** spécifie le code HTML qui s’affichera dans le volet Office du complément. Dans **Home.html**, remplacez l’élément `<body>` par le balisage suivant et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="25bd5-p102">**Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>

    ```html
    <body class="ms-font-m ms-welcome">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>Choose the button below to set the color of the selected range to green.</p>
                <br />
                <h3>Try it out</h3>
                <button class="ms-Button" id="set-color">Set color</button>
            </div>
        </div>
    </body>
    ```

2. <span data-ttu-id="25bd5-p103">Ouvrez le fichier **Home.js** à la racine du projet d’application web. Ce fichier spécifie le script pour le complément. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="25bd5-p103">Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.</span></span> 

    ```js
    'use strict';

    (function () {

        Office.onReady(function() {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                $('#set-color').click(setColor);
            });
        });

        function setColor() {
            Excel.run(function (context) {
                var range = context.workbook.getSelectedRange();
                range.format.fill.color = 'green';

                return context.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

3. <span data-ttu-id="25bd5-p104">Ouvrez le fichier **Home.css** à la racine du projet d’application web. Ce fichier spécifie les styles personnalisés pour le complément. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="25bd5-p104">Open the file **Home.css** in the root of the web application project. This file specifies the custom styles for the add-in. Replace the entire contents with the following code and save the file.</span></span> 

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

### <a name="update-the-manifest"></a><span data-ttu-id="25bd5-124">Mise à jour du manifeste</span><span class="sxs-lookup"><span data-stu-id="25bd5-124">Update the manifest</span></span>

1. <span data-ttu-id="25bd5-125">Ouvrez le fichier manifeste XML dans le projet de complément.</span><span class="sxs-lookup"><span data-stu-id="25bd5-125">Open the XML manifest file in the add-in project.</span></span> <span data-ttu-id="25bd5-126">Ce fichier définit les paramètres et les fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="25bd5-126">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="25bd5-p106">L’élément `ProviderName` possède une valeur d’espace réservé. Remplacez-le par votre nom.</span><span class="sxs-lookup"><span data-stu-id="25bd5-p106">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="25bd5-p107">L’attribut `DefaultValue` de l’élément `DisplayName` possède un espace réservé. Remplacez-le par **My Office Add-in**.</span><span class="sxs-lookup"><span data-stu-id="25bd5-p107">The `DefaultValue` attribute of the `DisplayName` element has a placeholder. Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="25bd5-p108">L’attribut `DefaultValue` de l’élément `Description` possède un espace réservé. Remplacez-le par **A task pane add-in for Excel**.</span><span class="sxs-lookup"><span data-stu-id="25bd5-p108">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for Excel**.</span></span>

5. <span data-ttu-id="25bd5-133">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="25bd5-133">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="25bd5-134">Essayez</span><span class="sxs-lookup"><span data-stu-id="25bd5-134">Try it out</span></span>

1. <span data-ttu-id="25bd5-p109">À l’aide de Visual Studio, testez le nouveau complément Excel en appuyant sur\*\* F5\*\* ou en choisissant le bouton **Démarrer** pour lancer Excel avec le bouton du complément \*\*Show Taskpane \*\*(Afficher le volet Office) qui apparaît dans le ruban. Le complément sera hébergé localement sur IIS.</span><span class="sxs-lookup"><span data-stu-id="25bd5-p109">Using Visual Studio, test the newly created Excel add-in by pressing **F5** or choosing the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="25bd5-137">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="25bd5-137">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Bouton Complément Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="25bd5-139">Sélectionnez une plage de cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="25bd5-139">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="25bd5-140">Dans le volet Office, cliquez sur le bouton **Définir couleur** pour définir la couleur de la plage sélectionnée en vert.</span><span class="sxs-lookup"><span data-stu-id="25bd5-140">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Complément Excel](../images/excel-quickstart-addin-2c.png)

# <a name="any-editortabvisual-studio-code"></a>[<span data-ttu-id="25bd5-142">Tous les éditeurs</span><span class="sxs-lookup"><span data-stu-id="25bd5-142">Any editor</span></span>](#tab/visual-studio-code)

### <a name="prerequisites"></a><span data-ttu-id="25bd5-143">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="25bd5-143">Prerequisites</span></span>

- [<span data-ttu-id="25bd5-144">Node.js</span><span class="sxs-lookup"><span data-stu-id="25bd5-144">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="25bd5-145">Installez la dernière version de [Yeoman](https://github.com/yeoman/yo) et le [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) globalement.</span><span class="sxs-lookup"><span data-stu-id="25bd5-145">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>
    ```command&nbsp;line
    npm install -g yo generator-office
    ```

### <a name="create-the-web-app"></a><span data-ttu-id="25bd5-146">Création de l’application web</span><span class="sxs-lookup"><span data-stu-id="25bd5-146">Create the web app</span></span>

1. <span data-ttu-id="25bd5-147">Utilisez le générateur Yeoman pour créer un projet de complément Excel.</span><span class="sxs-lookup"><span data-stu-id="25bd5-147">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="25bd5-148">Exécutez la commande suivante, puis répondez aux invites comme suit :</span><span class="sxs-lookup"><span data-stu-id="25bd5-148">Run the following command and then answer the prompts as follows:</span></span>

    ```command&nbsp;line
    yo office
    ```

    - <span data-ttu-id="25bd5-149">**Sélectionnez un type de projet :** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="25bd5-149">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="25bd5-150">**Sélectionnez un type de script :** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="25bd5-150">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="25bd5-151">**Comment souhaitez-vous nommer votre complément ? :** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="25bd5-151">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="25bd5-152">**Quelle application client Office voulez-vous prendre en charge ? :**`Excel`</span><span class="sxs-lookup"><span data-stu-id="25bd5-152">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Générateur Yeoman](../images/yo-office-jquery.png)

    <span data-ttu-id="25bd5-154">Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants de nœud de la prise en charge.</span><span class="sxs-lookup"><span data-stu-id="25bd5-154">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

2. <span data-ttu-id="25bd5-155">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="25bd5-155">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

### <a name="update-the-code"></a><span data-ttu-id="25bd5-156">Mise à jour du code</span><span class="sxs-lookup"><span data-stu-id="25bd5-156">Update the code</span></span> 

1. <span data-ttu-id="25bd5-p111">Dans votre éditeur de code, ouvrez **index.html** à la racine du projet. Ce fichier spécifie le code HTML qui s’affichera dans le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="25bd5-p111">In your code editor, open **index.html** in the root of the project. This file specifies the HTML that will be rendered in the add-in's task pane.</span></span> 

2. <span data-ttu-id="25bd5-159">Dans **index.html**, remplacez la balise `body` par le balisage suivant et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="25bd5-159">Within **index.html**, replace the `body` tag with the following markup and save the file.</span></span>

    ```html
    <body class="ms-font-m ms-welcome">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>Choose the button below to set the color of the selected range to green.</p>
                <br />
                <h3>Try it out</h3>
                <button class="ms-Button" id="set-color">Set color</button>
            </div>
        </div>
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>
    ```

3. <span data-ttu-id="25bd5-160">Ouvrez le fichier **src/index.js** pour spécifier le script pour le complément.</span><span class="sxs-lookup"><span data-stu-id="25bd5-160">Open the file **src\index.js** to specify the script for the add-in.</span></span> <span data-ttu-id="25bd5-161">Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="25bd5-161">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {
        Office.onReady(function() {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                $('#set-color').click(setColor);
            });
        });

        function setColor() {
            Excel.run(function (context) {
                var range = context.workbook.getSelectedRange();
                range.format.fill.color = 'green';

                return context.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

4. <span data-ttu-id="25bd5-p113">Ouvrez le fichier **app.css** pour spécifier les styles personnalisés pour le complément. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="25bd5-p113">Open the file **app.css** to specify the custom styles for the add-in. Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="25bd5-164">Mise à jour du manifeste</span><span class="sxs-lookup"><span data-stu-id="25bd5-164">Update the manifest</span></span>

1. <span data-ttu-id="25bd5-165">Ouvrez le fichier nommé **manifest.xml** pour définir les paramètres et les fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="25bd5-165">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="25bd5-p114">L’élément `ProviderName` possède une valeur d’espace réservé. Remplacez-le par votre nom.</span><span class="sxs-lookup"><span data-stu-id="25bd5-p114">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="25bd5-p115">L’attribut `DefaultValue` de l’élément `Description` possède un espace réservé. Remplacez-le par **A task pane add-in for Excel**.</span><span class="sxs-lookup"><span data-stu-id="25bd5-p115">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for Excel**.</span></span>

4. <span data-ttu-id="25bd5-170">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="25bd5-170">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="start-the-dev-server"></a><span data-ttu-id="25bd5-171">Démarrage du serveur de développement</span><span class="sxs-lookup"><span data-stu-id="25bd5-171">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a><span data-ttu-id="25bd5-172">Essayez !</span><span class="sxs-lookup"><span data-stu-id="25bd5-172">Try it out</span></span>

1. <span data-ttu-id="25bd5-173">Suivez les instructions pour la plateforme que vous utiliserez afin d’exécuter votre complément en vue d’en charger une version test dans Excel.</span><span class="sxs-lookup"><span data-stu-id="25bd5-173">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="25bd5-174">Windows : [Chargement de version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="25bd5-174">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="25bd5-175">Excel Online : [Chargement de versions test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="25bd5-175">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="25bd5-176">iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="25bd5-176">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="25bd5-177">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="25bd5-177">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Bouton Complément Excel](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="25bd5-179">Sélectionnez une plage de cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="25bd5-179">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="25bd5-180">Dans le volet Office, cliquez sur le bouton **Définir couleur** pour définir la couleur de la plage sélectionnée en vert.</span><span class="sxs-lookup"><span data-stu-id="25bd5-180">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Complément Excel](../images/excel-quickstart-addin-2c.png)

---

## <a name="next-steps"></a><span data-ttu-id="25bd5-182">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="25bd5-182">Next steps</span></span>

<span data-ttu-id="25bd5-p116">Félicitations, vous avez créé un complément Excel à l’aide de jQuery ! Découvrez à présent les fonctionnalités des compléments Excel et créez un complément plus complexe en continuant le didacticiel sur le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="25bd5-p116">Congratulations, you've successfully created an Excel add-in using jQuery! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="25bd5-185">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="25bd5-185">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="25bd5-186">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="25bd5-186">See also</span></span>

* [<span data-ttu-id="25bd5-187">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="25bd5-187">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="25bd5-188">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="25bd5-188">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="25bd5-189">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="25bd5-189">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="25bd5-190">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="25bd5-190">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
