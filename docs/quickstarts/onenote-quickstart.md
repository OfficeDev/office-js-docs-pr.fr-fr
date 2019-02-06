---
title: Créer votre premier complément OneNote
description: ''
ms.date: 01/17/2019
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: a0b2820f33e3a7cd31c12aec017ca552575a3f9b
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/05/2019
ms.locfileid: "29742337"
---
# <a name="build-your-first-onenote-add-in"></a><span data-ttu-id="f18e5-102">Créer votre premier complément OneNote</span><span class="sxs-lookup"><span data-stu-id="f18e5-102">Build your first OneNote add-in</span></span>

<span data-ttu-id="f18e5-103">Cet article décrit le processus de création d’un complément OneNote à l’aide de jQuery et de l’API JavaScript pour Office.</span><span class="sxs-lookup"><span data-stu-id="f18e5-103">In this article, you'll walk through the process of building a OneNote add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="f18e5-104">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="f18e5-104">Prerequisites</span></span>

- [<span data-ttu-id="f18e5-105">Node.js</span><span class="sxs-lookup"><span data-stu-id="f18e5-105">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="f18e5-106">Installez la dernière version de [Yeoman](https://github.com/yeoman/yo) et le [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) globalement.</span><span class="sxs-lookup"><span data-stu-id="f18e5-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in-project"></a><span data-ttu-id="f18e5-107">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="f18e5-107">Create the add-in project</span></span>

1. <span data-ttu-id="f18e5-108">Utilisez le générateur Yeoman afin de créer un projet de complément OneNote.</span><span class="sxs-lookup"><span data-stu-id="f18e5-108">Use the Yeoman generator to create a OneNote add-in project.</span></span> <span data-ttu-id="f18e5-109">Exécutez la commande suivante, puis répondez aux invites comme suit :</span><span class="sxs-lookup"><span data-stu-id="f18e5-109">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="f18e5-110">**Sélectionnez un type de projet :** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="f18e5-110">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="f18e5-111">**Sélectionnez un type de script :** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="f18e5-111">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="f18e5-112">**Comment souhaitez-vous nommer votre complément ? :** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="f18e5-112">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="f18e5-113">**Quelle application client Office voulez-vous prendre en charge ? :** `Onenote`</span><span class="sxs-lookup"><span data-stu-id="f18e5-113">**Which Office client application would you like to support?:** `Onenote`</span></span>

    ![Capture d’écran des invites et des réponses relatives au générateur Yeoman](../images/yo-office-onenote-jquery.png)
    
    <span data-ttu-id="f18e5-115">Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants de nœud de la prise en charge.</span><span class="sxs-lookup"><span data-stu-id="f18e5-115">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>
    
2. <span data-ttu-id="f18e5-116">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="f18e5-116">Navigate to the root folder of the project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="f18e5-117">Mise à jour du code</span><span class="sxs-lookup"><span data-stu-id="f18e5-117">Update the code</span></span>

1. <span data-ttu-id="f18e5-118">Dans votre éditeur de code, ouvrez **index.html** à la racine du projet.</span><span class="sxs-lookup"><span data-stu-id="f18e5-118">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="f18e5-119">Ce fichier contient le code HTML qui s’affichera dans le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="f18e5-119">This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="f18e5-120">Remplacez l’élément `<body>` par le balisage suivant et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="f18e5-120">Replace the `<body>` element with the following markup and save the file.</span></span> 

    ```html
    <body class="ms-font-m ms-welcome">
        <header class="ms-welcome__header ms-bgColor-themeDark ms-u-fadeIn500">
            <h2 class="ms-fontSize-xxl ms-fontWeight-regular ms-fontColor-white">OneNote Add-in</h1>
        </header>
        <main id="app-body" class="ms-welcome__main">
            <br />
            <p class="ms-font-m">Enter HTML content here:</p>
            <div class="ms-TextField ms-TextField--placeholder">
                <textarea id="textBox" rows="8" cols="30"></textarea>
            </div>
            <button id="addOutline" class="ms-Button ms-Button--primary">
                <span class="ms-Button-label">Add outline</span>
            </button>
        </main>
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>
    ```

3. <span data-ttu-id="f18e5-121">Ouvrez le fichier **src/index.js** afin de spécifier le script pour le complément.</span><span class="sxs-lookup"><span data-stu-id="f18e5-121">Open the file **src\index.js** to specify the script for the add-in.</span></span> <span data-ttu-id="f18e5-122">Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="f18e5-122">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    import * as OfficeHelpers from "@microsoft/office-js-helpers";

    Office.onReady(() => {
        // Office is ready
        $(document).ready(() => {
            // The document is ready
            $('#addOutline').click(addOutlineToPage);
        });
    });
    
    async function addOutlineToPage() {
        try {
            await OneNote.run(async context => {
                var html = "<p>" + $("#textBox").val() + "</p>";

                // Get the current page.
                var page = context.application.getActivePage();

                // Queue a command to load the page with the title property.
                page.load("title");

                // Add text to the page by using the specified HTML.
                var outline = page.addOutline(40, 90, html);

                // Run the queued commands, and return a promise to indicate task completion.
                return context.sync()
                    .then(function() {
                        console.log("Added outline to page " + page.title);
                    })
                    .catch(function(error) {
                        app.showNotification("Error: " + error);
                        console.log("Error: " + error);
                        if (error instanceof OfficeExtension.Error) {
                            console.log("Debug info: " + JSON.stringify(error.debugInfo));
                        }
                    });
                });
        } catch (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        }
    }
    ```

4. <span data-ttu-id="f18e5-123">Ouvrez le fichier **app.css** pour spécifier les styles personnalisés pour le complément.</span><span class="sxs-lookup"><span data-stu-id="f18e5-123">Open the file **app.css** to specify the custom styles for the add-in.</span></span> <span data-ttu-id="f18e5-124">Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="f18e5-124">Replace the entire contents with the following and save the file.</span></span>

    ```css
    html, body {
        width: 100%;
        height: 100%;
        margin: 0;
        padding: 0;
    }

    ul, p, h1, h2, h3, h4, h5, h6 {
        margin: 0;
        padding: 0;
    }

    .ms-welcome {
        position: relative;
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        min-height: 500px;
        min-width: 320px;
        overflow: auto;
        overflow-x: hidden;
    }

    .ms-welcome__header {
        min-height: 30px;
        padding: 0px;
        padding-bottom: 5px;
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        -webkit-align-items: center;
        align-items: center;
        -webkit-justify-content: flex-end;
        justify-content: flex-end;
    }

    .ms-welcome__header > h1 {
        margin-top: 5px;
        text-align: center;
    }

    .ms-welcome__main {
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        -webkit-align-items: center;
        align-items: left;
        -webkit-flex: 1 0 0;
        flex: 1 0 0;
        padding: 30px 20px;
    }

    .ms-welcome__main > h2 {
        width: 100%;
        text-align: left;
    }

    @media (min-width: 0) and (max-width: 350px) {
        .ms-welcome__features {
            width: 100%;
        }
    }
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="f18e5-125">Mise à jour du manifeste</span><span class="sxs-lookup"><span data-stu-id="f18e5-125">Update the manifest</span></span>

1. <span data-ttu-id="f18e5-126">Ouvrez le fichier nommé **manifest.xml** pour définir les paramètres et les fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="f18e5-126">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="f18e5-127">L’élément `ProviderName` possède une valeur d’espace réservé.</span><span class="sxs-lookup"><span data-stu-id="f18e5-127">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="f18e5-128">Remplacez-le par votre nom.</span><span class="sxs-lookup"><span data-stu-id="f18e5-128">Replace it with your name.</span></span>

3. <span data-ttu-id="f18e5-129">L’attribut `DefaultValue` de l’élément `Description` possède un espace réservé.</span><span class="sxs-lookup"><span data-stu-id="f18e5-129">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="f18e5-130">Remplacez-le par **A task pane add-in for OneNote**.</span><span class="sxs-lookup"><span data-stu-id="f18e5-130">Replace it with **A task pane add-in for OneNote**.</span></span>

4. <span data-ttu-id="f18e5-131">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="f18e5-131">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for OneNote"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="f18e5-132">Démarrage du serveur de développement</span><span class="sxs-lookup"><span data-stu-id="f18e5-132">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)]

## <a name="try-it-out"></a><span data-ttu-id="f18e5-133">Essayez !</span><span class="sxs-lookup"><span data-stu-id="f18e5-133">Try it out</span></span>

1. <span data-ttu-id="f18e5-134">Dans [OneNote Online](https://www.onenote.com/notebooks), ouvrez un bloc-notes.</span><span class="sxs-lookup"><span data-stu-id="f18e5-134">In [OneNote Online](https://www.onenote.com/notebooks), open a notebook.</span></span>

2. <span data-ttu-id="f18e5-135">Choisissez **Insertion > Compléments Office** pour ouvrir la boîte de dialogue Compléments Office.</span><span class="sxs-lookup"><span data-stu-id="f18e5-135">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="f18e5-136">Si vous êtes connecté avec votre compte de consommateur, sélectionnez l’onglet **MES COMPLÉMENTS**, puis choisissez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="f18e5-136">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="f18e5-137">Si vous êtes connecté avec votre compte professionnel ou scolaire, sélectionnez l’onglet **MON ORGANISATION**, puis choisissez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="f18e5-137">If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span> 

    <span data-ttu-id="f18e5-138">L’image suivante montre l’onglet **MES COMPLÉMENTS** pour les blocs-notes de consommateurs.</span><span class="sxs-lookup"><span data-stu-id="f18e5-138">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. <span data-ttu-id="f18e5-139">Dans la boîte de dialogue Télécharger le complément, accédez à **manifest.xml** dans le dossier de projet, puis choisissez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="f18e5-139">In the Upload Add-in dialog, browse to **manifest.xml** in your project folder, and then choose **Upload**.</span></span> 

4. <span data-ttu-id="f18e5-140">Dans l’onglet **Accueil**, choisissez le bouton **Afficher le volet de tâches** du ruban.</span><span class="sxs-lookup"><span data-stu-id="f18e5-140">From the **Home** tab, choose the **Show Taskpane** button in the ribbon.</span></span> <span data-ttu-id="f18e5-141">Le volet Office du complément s’ouvre dans un iFrame à côté de la page OneNote.</span><span class="sxs-lookup"><span data-stu-id="f18e5-141">The add-in task pane opens in an iFrame next to the OneNote page.</span></span>

5. <span data-ttu-id="f18e5-142">Entrez le contenu HTML suivant dans la zone de texte, puis sélectionnez **Ajouter un plan**.</span><span class="sxs-lookup"><span data-stu-id="f18e5-142">Enter the following HTML content in the text area, and then choose **Add outline**.</span></span>  

    ```html
    <ol>
    <li>Item #1</li>
    <li>Item #2</li>
    <li>Item #3</li>
    <li>Item #4</li>
    </ol>
    ```

    <span data-ttu-id="f18e5-143">Le plan que vous avez spécifié est ajouté à la page.</span><span class="sxs-lookup"><span data-stu-id="f18e5-143">The outline that you specified is added to the page.</span></span>

    ![Complément OneNote généré à partir de cette procédure pas à pas](../images/onenote-first-add-in-3.png)

## <a name="troubleshooting-and-tips"></a><span data-ttu-id="f18e5-145">Conseils et résolution des problèmes</span><span class="sxs-lookup"><span data-stu-id="f18e5-145">Troubleshooting and tips</span></span>

- <span data-ttu-id="f18e5-p108">Vous pouvez déboguer le complément à l’aide des outils de développement de votre navigateur. Lorsque vous utilisez le serveur web Gulp et le débogage dans Internet Explorer ou Chrome, vous pouvez enregistrer les modifications localement et simplement actualiser l’iFrame du complément.</span><span class="sxs-lookup"><span data-stu-id="f18e5-p108">You can debug the add-in using your browser's developer tools. When you're using the Gulp web server and debugging in Internet Explorer or Chrome, you can save your changes locally and then just refresh the add-in's iFrame.</span></span>

- <span data-ttu-id="f18e5-p109">Lorsque vous examinez un objet OneNote, les propriétés qui sont actuellement disponibles affichent les valeurs réelles. Les propriétés qui doivent être chargées sont affichées comme *non définies*. Développez le nœud `_proto_` pour visualiser les propriétés qui sont définies sur l’objet, mais qui ne sont pas encore chargées.</span><span class="sxs-lookup"><span data-stu-id="f18e5-p109">When you inspect a OneNote object, the properties that are currently available for use display actual values. Properties that need to be loaded display *undefined*. Expand the `_proto_` node to see properties that are defined on the object but are not yet loaded.</span></span>

   ![Objet OneNote déchargé dans le débogueur](../images/onenote-debug.png)

- <span data-ttu-id="f18e5-p110">Vous devez activer le contenu mixte dans le navigateur si votre complément utilise des ressources HTTP. Les compléments de production doivent uniquement utiliser des ressources HTTPS sécurisées.</span><span class="sxs-lookup"><span data-stu-id="f18e5-p110">You need to enable mixed content in the browser if your add-in uses any HTTP resources. Production add-ins should use only secure HTTPS resources.</span></span>

- <span data-ttu-id="f18e5-154">Les compléments de volet Office peuvent être ouverts à partir de n’importe où, mais les compléments de contenu peuvent uniquement être insérés à l’intérieur de contenu de page normal (et non dans des titres, des images, des iFrames, etc.).</span><span class="sxs-lookup"><span data-stu-id="f18e5-154">Task pane add-ins can be opened from anywhere, but content add-ins can only be inserted inside regular page content (i.e. not in titles, images, iFrames, etc.).</span></span> 

## <a name="next-steps"></a><span data-ttu-id="f18e5-155">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="f18e5-155">Next steps</span></span>

<span data-ttu-id="f18e5-156">Félicitations, vous avez créé un complément OneNote !</span><span class="sxs-lookup"><span data-stu-id="f18e5-156">Congratulations, you've successfully created a OneNote add-in!</span></span> <span data-ttu-id="f18e5-157">Ensuite, vous allez étudier en détail les concepts fondamentaux de la création de compléments Excel.</span><span class="sxs-lookup"><span data-stu-id="f18e5-157">Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="f18e5-158">Vue d’ensemble de la programmation de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="f18e5-158">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="f18e5-159">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f18e5-159">See also</span></span>

- [<span data-ttu-id="f18e5-160">Vue d’ensemble de la programmation de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="f18e5-160">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="f18e5-161">Référence de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="f18e5-161">OneNote JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="f18e5-162">Exemple de grille d’évaluation</span><span class="sxs-lookup"><span data-stu-id="f18e5-162">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="f18e5-163">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="f18e5-163">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)

