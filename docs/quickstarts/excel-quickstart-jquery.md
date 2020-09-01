---
title: Créer votre premier complément du volet des tâches d’Excel
description: Découvrez comment créer un complément de volet des tâches Excel simple à l’aide de l’API JavaScript pour Office.
ms.date: 04/03/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 4043fa56d844ca1160c61dd94d229172682c3af2
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292340"
---
# <a name="build-an-excel-task-pane-add-in"></a><span data-ttu-id="4480d-103">Créer un complément de volet de tâches Excel</span><span class="sxs-lookup"><span data-stu-id="4480d-103">Build an Excel task pane add-in</span></span>

<span data-ttu-id="4480d-104">Dans cet article, vous découvrirez comment créer un complément de volet de tâches Excel.</span><span class="sxs-lookup"><span data-stu-id="4480d-104">In this article, you'll walk through the process of building an Excel task pane add-in.</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="4480d-105">Créer le complément</span><span class="sxs-lookup"><span data-stu-id="4480d-105">Create the add-in</span></span> 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]
# <a name="yeoman-generator"></a>[<span data-ttu-id="4480d-106">Générateur Yeoman</span><span class="sxs-lookup"><span data-stu-id="4480d-106">Yeoman generator</span></span>](#tab/yeomangenerator)

[!include[Redirect to the single sign-on (SSO) quick start](../includes/sso-quickstart-reference.md)]

## <a name="prerequisites"></a><span data-ttu-id="4480d-107">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="4480d-107">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="4480d-108">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="4480d-108">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="4480d-109">**Sélectionnez un type de projet :** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="4480d-109">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="4480d-110">**Sélectionnez un type de script :** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="4480d-110">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="4480d-111">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="4480d-111">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="4480d-112">**Quelle application client Office voulez-vous prendre en charge ?**</span><span class="sxs-lookup"><span data-stu-id="4480d-112">**Which Office client application would you like to support?**</span></span> `Excel`

![Générateur Yeoman](../images/yo-office-excel.png)

<span data-ttu-id="4480d-114">Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="4480d-114">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="explore-the-project"></a><span data-ttu-id="4480d-115">Explorer le projet</span><span class="sxs-lookup"><span data-stu-id="4480d-115">Explore the project</span></span>

[!include[Yeoman generator add-in project components](../includes/yo-task-pane-project-components-js.md)]

### <a name="try-it-out"></a><span data-ttu-id="4480d-116">Essayez</span><span class="sxs-lookup"><span data-stu-id="4480d-116">Try it out</span></span>

1. <span data-ttu-id="4480d-117">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="4480d-117">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)]

3. <span data-ttu-id="4480d-118">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="4480d-118">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Bouton Complément Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="4480d-120">Sélectionnez une plage de cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="4480d-120">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="4480d-121">En bas du volet Office, cliquez sélectionnez le lien **Exécuter** pour définir la couleur de la plage sélectionnée sur jaune.</span><span class="sxs-lookup"><span data-stu-id="4480d-121">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Complément Excel avec le bouton Exécuter](../images/excel-quickstart-addin-3c.png)

### <a name="next-steps"></a><span data-ttu-id="4480d-123">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="4480d-123">Next steps</span></span>

<span data-ttu-id="4480d-124">Félicitations, vous avez créé un complément de volet de tâches Excel !</span><span class="sxs-lookup"><span data-stu-id="4480d-124">Congratulations, you've successfully created an Excel task pane add-in!</span></span> <span data-ttu-id="4480d-125">Ensuite, découvrez les fonctionnalités d’un complément Excel et créez-en un plus complexe en suivant le [didacticiel sur les compléments Excel](../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="4480d-125">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the [Excel add-in tutorial](../tutorials/excel-tutorial.md).</span></span>

# <a name="visual-studio"></a>[<span data-ttu-id="4480d-126">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="4480d-126">Visual Studio</span></span>](#tab/visualstudio)

### <a name="prerequisites"></a><span data-ttu-id="4480d-127">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="4480d-127">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="4480d-128">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="4480d-128">Create the add-in project</span></span>

1. <span data-ttu-id="4480d-129">Dans Visual Studio, choisissez **Créer un nouveau projet**.</span><span class="sxs-lookup"><span data-stu-id="4480d-129">In Visual Studio, choose **Create a new project**.</span></span>

2. <span data-ttu-id="4480d-130">À l’aide de la zone de recherche, entrez **complément**.</span><span class="sxs-lookup"><span data-stu-id="4480d-130">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="4480d-131">Choisissez **Complément web Excel**, puis sélectionnez **Suivant**.</span><span class="sxs-lookup"><span data-stu-id="4480d-131">Choose **Excel Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="4480d-132">Nommez votre projet et sélectionnez **Créer**.</span><span class="sxs-lookup"><span data-stu-id="4480d-132">Name your project and select **Create**.</span></span>

4. <span data-ttu-id="4480d-133">Dans la fenêtre de dialogue **Créer un complément Office**, sélectionnez **Ajouter de nouvelles fonctionnalités à Excel**, puis sélectionnez **Terminer** pour créer le projet.</span><span class="sxs-lookup"><span data-stu-id="4480d-133">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="4480d-p103">Visual Studio crée une solution et ses deux projets apparaissent dans l’**explorateur de solutions**. Le fichier **Home.html** s’ouvre dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="4480d-p103">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="4480d-136">Explorer la solution Visual Studio</span><span class="sxs-lookup"><span data-stu-id="4480d-136">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="4480d-137">Mise à jour du code</span><span class="sxs-lookup"><span data-stu-id="4480d-137">Update the code</span></span>

1. <span data-ttu-id="4480d-p104">**Home.html** spécifie le code HTML qui s’affichera dans le volet Office du complément. Dans **Home.html**, remplacez l’élément `<body>` par le balisage suivant et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="4480d-p104">**Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>

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

2. <span data-ttu-id="4480d-p105">Ouvrez le fichier **Home.js** à la racine du projet d’application web. Ce fichier spécifie le script pour le complément. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="4480d-p105">Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.</span></span>

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

3. <span data-ttu-id="4480d-p106">Ouvrez le fichier **Home.css** à la racine du projet d’application web. Ce fichier spécifie les styles personnalisés pour le complément. Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="4480d-p106">Open the file **Home.css** in the root of the web application project. This file specifies the custom styles for the add-in. Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="4480d-146">Mise à jour du manifeste</span><span class="sxs-lookup"><span data-stu-id="4480d-146">Update the manifest</span></span>

1. <span data-ttu-id="4480d-147">Ouvrez le fichier manifeste XML dans le projet de complément.</span><span class="sxs-lookup"><span data-stu-id="4480d-147">Open the XML manifest file in the add-in project.</span></span> <span data-ttu-id="4480d-148">Ce fichier définit les paramètres et les fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="4480d-148">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="4480d-p108">L’élément `ProviderName` possède une valeur d’espace réservé. Remplacez-le par votre nom.</span><span class="sxs-lookup"><span data-stu-id="4480d-p108">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="4480d-p109">L’attribut `DefaultValue` de l’élément `DisplayName` possède un espace réservé. Remplacez-le par **My Office Add-in**.</span><span class="sxs-lookup"><span data-stu-id="4480d-p109">The `DefaultValue` attribute of the `DisplayName` element has a placeholder. Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="4480d-p110">L’attribut `DefaultValue` de l’élément `Description` possède un espace réservé. Remplacez-le par **A task pane add-in for Excel**.</span><span class="sxs-lookup"><span data-stu-id="4480d-p110">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for Excel**.</span></span>

5. <span data-ttu-id="4480d-155">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="4480d-155">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="4480d-156">Essayez</span><span class="sxs-lookup"><span data-stu-id="4480d-156">Try it out</span></span>

1. <span data-ttu-id="4480d-157">À l’aide de Visual Studio, testez le nouveau complément Excel en appuyant sur **F5** ou en choisissant le bouton **Démarrer** pour lancer Excel avec le bouton du complément **Afficher le volet Office** qui apparaît dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="4480d-157">Using Visual Studio, test the newly created Excel add-in by pressing **F5** or choosing the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="4480d-158">Le complément est hébergé localement sur IIS.</span><span class="sxs-lookup"><span data-stu-id="4480d-158">The add-in will be hosted locally on IIS.</span></span> <span data-ttu-id="4480d-159">Si on vous demande d’approuver un certificat, faites-le pour autoriser le complément à se connecter à son application Office.</span><span class="sxs-lookup"><span data-stu-id="4480d-159">If you are asked to trust a certificate, do so to allow the add-in to connect to its Office application.</span></span>

2. <span data-ttu-id="4480d-160">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="4480d-160">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Bouton du complément Excel Afficher le volet de tâches](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="4480d-162">Sélectionnez une plage de cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="4480d-162">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="4480d-163">Dans le volet Office, cliquez sur le bouton **Définir couleur** pour définir la couleur de la plage sélectionnée en vert.</span><span class="sxs-lookup"><span data-stu-id="4480d-163">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Complément Excel](../images/excel-quickstart-addin-2c.png)

[!include[Console tool note](../includes/console-tool-note.md)]

### <a name="next-steps"></a><span data-ttu-id="4480d-165">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="4480d-165">Next steps</span></span>

<span data-ttu-id="4480d-166">Félicitations, vous avez créé un complément de volet de tâches Excel !</span><span class="sxs-lookup"><span data-stu-id="4480d-166">Congratulations, you've successfully created an Excel task pane add-in!</span></span> <span data-ttu-id="4480d-167">Ensuite, en savoir plus sur la [création de compléments Office avec Visual Studio](../develop/develop-add-ins-visual-studio.md).</span><span class="sxs-lookup"><span data-stu-id="4480d-167">Next, learn more about [developing Office Add-ins with Visual Studio](../develop/develop-add-ins-visual-studio.md).</span></span>

---

## <a name="see-also"></a><span data-ttu-id="4480d-168">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="4480d-168">See also</span></span>

* [<span data-ttu-id="4480d-169">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="4480d-169">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="4480d-170">Création de compléments Office</span><span class="sxs-lookup"><span data-stu-id="4480d-170">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="4480d-171">Développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="4480d-171">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="4480d-172">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="4480d-172">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="4480d-173">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="4480d-173">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="4480d-174">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="4480d-174">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
