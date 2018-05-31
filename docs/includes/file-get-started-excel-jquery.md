# <a name="build-an-excel-add-in-using-jquery"></a><span data-ttu-id="60858-101">Développement d’un complément Excel à l’aide de jQuery</span><span class="sxs-lookup"><span data-stu-id="60858-101">Build an Excel add-in using jQuery</span></span>

<span data-ttu-id="60858-102">Cet article décrit le processus de création d’un complément Excel à l’aide de jQuery et de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="60858-102">In this article, you'll walk through the process of building an Excel add-in by using jQuery and the Excel JavaScript API.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="60858-103">Créer le complément</span><span class="sxs-lookup"><span data-stu-id="60858-103">Create the add-in</span></span> 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[<span data-ttu-id="60858-104">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="60858-104">Visual Studio</span></span>](#tab/visual-studio)

### <a name="prerequisites"></a><span data-ttu-id="60858-105">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="60858-105">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="60858-106">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="60858-106">Create the add-in project</span></span>

1. <span data-ttu-id="60858-107">Dans la barre de menu de Visual Studio, choisissez successivement **Fichier** > **Nouveau** > **Projet**.</span><span class="sxs-lookup"><span data-stu-id="60858-107">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="60858-108">Dans la liste des types de projet, sous **Visual C#** ou **Visual Basic**, développez **Office/SharePoint**, choisissez **Compléments**, puis **Complément Excel Web** pour le type de projet.</span><span class="sxs-lookup"><span data-stu-id="60858-108">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="60858-109">Nommez le projet, puis cliquez sur **OK**.</span><span class="sxs-lookup"><span data-stu-id="60858-109">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="60858-110">Dans la fenêtre de dialogue **Créer un complément Office**, sélectionnez **Ajouter de nouvelles fonctionnalités à Excel**, puis sélectionnez **Terminer** pour créer le projet.</span><span class="sxs-lookup"><span data-stu-id="60858-110">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="60858-p101">Visual Studio crée une solution et ses deux projets apparaissent dans l’**explorateur de solutions**. Le fichier **Home.html** s’ouvre dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="60858-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>
    
### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="60858-113">Explorer la solution Visual Studio</span><span class="sxs-lookup"><span data-stu-id="60858-113">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="60858-114">Mise à jour du code</span><span class="sxs-lookup"><span data-stu-id="60858-114">Update the code</span></span>

1. <span data-ttu-id="60858-115">**Home.html** spécifie le code HTML qui s’affichera dans le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="60858-115">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="60858-116">Dans **Home.html**, remplacez l’élément `<body>` par le balisage suivant et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="60858-116">In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
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

2. <span data-ttu-id="60858-117">Ouvrez le fichier **Home.js** à la racine du projet d’application web.</span><span class="sxs-lookup"><span data-stu-id="60858-117">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="60858-118">Ce fichier spécifie le script pour le complément.</span><span class="sxs-lookup"><span data-stu-id="60858-118">This file specifies the script for the add-in.</span></span> <span data-ttu-id="60858-119">Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="60858-119">Replace the entire contents with the following code and save the file.</span></span> 

    ```js
    'use strict';

    (function () {
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#set-color').click(setColor);
            });
        };

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

3. <span data-ttu-id="60858-120">Ouvrez le fichier **Home.css** à la racine du projet d’application web.</span><span class="sxs-lookup"><span data-stu-id="60858-120">Open the file **Home.css** in the root of the web application project.</span></span> <span data-ttu-id="60858-121">Ce fichier spécifie les styles personnalisés pour le complément.</span><span class="sxs-lookup"><span data-stu-id="60858-121">This file specifies the custom styles for the add-in.</span></span> <span data-ttu-id="60858-122">Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="60858-122">Replace the entire contents with the following code and save the file.</span></span> 

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

### <a name="update-the-manifest"></a><span data-ttu-id="60858-123">Mise à jour du manifeste</span><span class="sxs-lookup"><span data-stu-id="60858-123">Update the manifest</span></span>

1. <span data-ttu-id="60858-124">Ouvrez le fichier manifeste XML dans le projet de complément.</span><span class="sxs-lookup"><span data-stu-id="60858-124">Open the XML manifest file in the Add-in project.</span></span> <span data-ttu-id="60858-125">Ce fichier définit les paramètres et les fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="60858-125">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="60858-126">L’élément `ProviderName` possède une valeur d’espace réservé.</span><span class="sxs-lookup"><span data-stu-id="60858-126">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="60858-127">Remplacez-le par votre nom.</span><span class="sxs-lookup"><span data-stu-id="60858-127">Replace it with your name.</span></span>

3. <span data-ttu-id="60858-128">L’attribut `DefaultValue` de l’élément `DisplayName` possède un espace réservé.</span><span class="sxs-lookup"><span data-stu-id="60858-128">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="60858-129">Remplacez-le par **My Office Add-in**.</span><span class="sxs-lookup"><span data-stu-id="60858-129">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="60858-130">L’attribut `DefaultValue` de l’élément `Description` possède un espace réservé.</span><span class="sxs-lookup"><span data-stu-id="60858-130">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="60858-131">Remplacez-le par **A task pane add-in for Excel**.</span><span class="sxs-lookup"><span data-stu-id="60858-131">Replace it with **A task pane add-in for Excel**.</span></span>

5. <span data-ttu-id="60858-132">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="60858-132">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="60858-133">Essayez !</span><span class="sxs-lookup"><span data-stu-id="60858-133">Try it out</span></span>

1. <span data-ttu-id="60858-p109">À l’aide de Visual Studio, testez le nouveau complément Excel en appuyant sur F5 ou en choisissant le bouton **Démarrer** pour lancer Excel avec le bouton du complément **Show Taskpane** (Afficher le volet Office) qui apparaît dans le ruban. Le complément sera hébergé localement sur IIS.</span><span class="sxs-lookup"><span data-stu-id="60858-p109">Using Visual Studio, test the newly created Excel add-in by pressing F5 or choosing the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="60858-136">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="60858-136">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Bouton Complément Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="60858-138">Sélectionnez une plage de cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="60858-138">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="60858-139">Dans le volet Office, cliquez sur le bouton **Définir couleur** pour définir la couleur de la plage sélectionnée en vert.</span><span class="sxs-lookup"><span data-stu-id="60858-139">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Complément Excel](../images/excel-quickstart-addin-2c.png)

# <a name="any-editortabvisual-studio-code"></a>[<span data-ttu-id="60858-141">Tous les éditeurs</span><span class="sxs-lookup"><span data-stu-id="60858-141">Any editor</span></span>](#tab/visual-studio-code)

### <a name="prerequisites"></a><span data-ttu-id="60858-142">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="60858-142">Prerequisites</span></span>

- [<span data-ttu-id="60858-143">Node.js</span><span class="sxs-lookup"><span data-stu-id="60858-143">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="60858-144">Installez la dernière version de [Yeoman](https://github.com/yeoman/yo) et le [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) globalement.</span><span class="sxs-lookup"><span data-stu-id="60858-144">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-web-app"></a><span data-ttu-id="60858-145">Création de l’application web</span><span class="sxs-lookup"><span data-stu-id="60858-145">Create the web app</span></span>

1. <span data-ttu-id="60858-146">Créez un dossier sur votre lecteur local et nommez-le **my-addin**.</span><span class="sxs-lookup"><span data-stu-id="60858-146">Create a folder on your local drive and name it **my-addin**.</span></span> <span data-ttu-id="60858-147">Il s’agit de l’endroit où vous allez créer les fichiers de votre application.</span><span class="sxs-lookup"><span data-stu-id="60858-147">This is where you'll create the files for your app.</span></span>

2. <span data-ttu-id="60858-148">Accédez au dossier de votre application.</span><span class="sxs-lookup"><span data-stu-id="60858-148">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

3. <span data-ttu-id="60858-149">Utilisez le générateur Yeoman pour générer le fichier manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="60858-149">Use the Yeoman generator to generate the manifest file for your add-in.</span></span> <span data-ttu-id="60858-150">Exécutez la commande suivante, puis répondez aux invites comme indiqué dans la capture d’écran suivante :</span><span class="sxs-lookup"><span data-stu-id="60858-150">Run the following command and then answer the prompts as shown in the following screenshot:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="60858-151">**Voulez-vous créer un sous-dossier de votre projet ? :** `No`</span><span class="sxs-lookup"><span data-stu-id="60858-151">**Would you like to create a new subfolder for your project?:** `No`</span></span>
    - <span data-ttu-id="60858-152">**Comment souhaitez-vous nommer votre complément ? :** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="60858-152">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="60858-153">**Quelle application client Office voulez-vous prendre en charge ? :** `Excel`</span><span class="sxs-lookup"><span data-stu-id="60858-153">**Which Office client application would you like to support?:** `Excel`</span></span>
    - <span data-ttu-id="60858-154">**Voulez-vous créer un complément ? :** `Yes`</span><span class="sxs-lookup"><span data-stu-id="60858-154">**Would you like to create a new add-in?:** `Yes`</span></span>
    - <span data-ttu-id="60858-155">**Souhaitez-vous utiliser TypeScript ? :** `No`</span><span class="sxs-lookup"><span data-stu-id="60858-155">**Would you like to use TypeScript?:** `No`</span></span>
    - <span data-ttu-id="60858-156">**Choisissez une infrastructure :** `Jquery`</span><span class="sxs-lookup"><span data-stu-id="60858-156">**Choose a framework:** `Jquery`</span></span>

    <span data-ttu-id="60858-p112">Le générateur demande ensuite si vous voulez ouvrir **resource.html**. Il n’est pas nécessaire de l’ouvrir pour ce didacticiel, mais n’hésitez pas à l’ouvrir si vous êtes curieux. Cliquez sur Oui ou Non pour fermer l’assistant et laisser le générateur faire son travail.</span><span class="sxs-lookup"><span data-stu-id="60858-p112">The generator will then ask you if you want to open **resource.html**. It isn't necessary to open it for this tutorial, but feel free to open it if you're curious! Choose yes or no to complete the wizard and allow the generator to do its work.</span></span>

    ![Générateur Yeoman](../images/yo-office-jquery.png)


4. <span data-ttu-id="60858-161">Dans votre éditeur de code, ouvrez **index.html** à la racine du projet.</span><span class="sxs-lookup"><span data-stu-id="60858-161">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="60858-162">Ce fichier spécifie le code HTML qui s’affichera dans le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="60858-162">This file specifies the HTML that will be rendered in the add-in's task pane.</span></span> 
 
5. <span data-ttu-id="60858-163">Dans **index.html**, remplacez la balise `header` générée par le balisage suivant.</span><span class="sxs-lookup"><span data-stu-id="60858-163">Within **index.html**, replace the generated `header` tag with the following markup.</span></span>
 
    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    ```

6. <span data-ttu-id="60858-164">Dans **index.html**, remplacez la balise `main` générée par le balisage suivant et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="60858-164">Within **index.html**, replace the generated `main` tag with the following markup, and save the file.</span></span>

    ```html
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <button class="ms-Button" id="set-color">Set color</button>
        </div>
    </div>
    ```

7. <span data-ttu-id="60858-165">Ouvrez le fichier **app.js** pour spécifier le script pour le complément.</span><span class="sxs-lookup"><span data-stu-id="60858-165">Open the file **app.js** to specify the script for the add-in.</span></span> <span data-ttu-id="60858-166">Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="60858-166">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';
    
    (function () {
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#set-color').click(setColor);
            });
        };

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

8. <span data-ttu-id="60858-167">Ouvrez le fichier **app.css** pour spécifier les styles personnalisés pour le complément.</span><span class="sxs-lookup"><span data-stu-id="60858-167">Open the file **app.css** to specify the custom styles for the add-in.</span></span> <span data-ttu-id="60858-168">Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="60858-168">Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="60858-169">Mise à jour du manifeste</span><span class="sxs-lookup"><span data-stu-id="60858-169">Update the manifest</span></span>

1. <span data-ttu-id="60858-170">Ouvrez le fichier nommé **my-office-add-in-manifest.xml** pour définir les paramètres et les fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="60858-170">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span> 

2. <span data-ttu-id="60858-171">L’élément `ProviderName` possède une valeur d’espace réservé.</span><span class="sxs-lookup"><span data-stu-id="60858-171">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="60858-172">Remplacez-le par votre nom.</span><span class="sxs-lookup"><span data-stu-id="60858-172">Replace it with your name.</span></span>

3. <span data-ttu-id="60858-173">L’attribut `DefaultValue` de l’élément `DisplayName` possède un espace réservé.</span><span class="sxs-lookup"><span data-stu-id="60858-173">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="60858-174">Remplacez-le par **My Office Add-in**.</span><span class="sxs-lookup"><span data-stu-id="60858-174">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="60858-175">L’attribut `DefaultValue` de l’élément `Description` possède un espace réservé.</span><span class="sxs-lookup"><span data-stu-id="60858-175">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="60858-176">Remplacez-le par **A task pane add-in for Excel**.</span><span class="sxs-lookup"><span data-stu-id="60858-176">Replace it with **A task pane add-in for Excel**.</span></span>

5. <span data-ttu-id="60858-177">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="60858-177">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="start-the-dev-server"></a><span data-ttu-id="60858-178">Démarrage du serveur de développement</span><span class="sxs-lookup"><span data-stu-id="60858-178">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a><span data-ttu-id="60858-179">Essayez !</span><span class="sxs-lookup"><span data-stu-id="60858-179">Try it out</span></span>

1. <span data-ttu-id="60858-180">Suivez les instructions pour la plateforme que vous utiliserez afin d’exécuter votre complément en vue d’en charger une version test dans Excel.</span><span class="sxs-lookup"><span data-stu-id="60858-180">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="60858-181">Windows : [Chargement de version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="60858-181">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="60858-182">Excel Online : [Chargement de versions test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="60858-182">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="60858-183">iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="60858-183">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="60858-184">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="60858-184">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Bouton Complément Excel](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="60858-186">Sélectionnez une plage de cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="60858-186">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="60858-187">Dans le volet Office, cliquez sur le bouton **Définir couleur** pour définir la couleur de la plage sélectionnée en vert.</span><span class="sxs-lookup"><span data-stu-id="60858-187">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Complément Excel](../images/excel-quickstart-addin-2c.png)

---

## <a name="next-steps"></a><span data-ttu-id="60858-189">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="60858-189">Next steps</span></span>

<span data-ttu-id="60858-p119">Félicitations, vous avez créé un complément Excel à l’aide de jQuery ! Découvrez à présent les fonctionnalités des compléments Excel et créez un complément plus complexe en continuant le didacticiel sur le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="60858-p119">Congratulations, you've successfully created an Excel add-in using jQuery! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="60858-192">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="60858-192">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="60858-193">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="60858-193">See also</span></span>

* [<span data-ttu-id="60858-194">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="60858-194">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="60858-195">Concepts de base de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="60858-195">Excel JavaScript API core concepts</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="60858-196">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="60858-196">Excel add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="60858-197">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="60858-197">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
