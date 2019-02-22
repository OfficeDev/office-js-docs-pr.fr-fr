---
title: Créer un complément Excel à l’aide d’Angular
description: ''
ms.date: 10/19/2018
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: da3e0ce9c6cb075ff828344956cfe25fb579a8fa
ms.sourcegitcommit: 03773fef3d2a380028ba0804739d2241d4b320e5
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2019
ms.locfileid: "30091252"
---
# <a name="build-an-excel-add-in-using-angular"></a><span data-ttu-id="96768-102">Créer un complément Excel à l’aide d’Angular</span><span class="sxs-lookup"><span data-stu-id="96768-102">Build an Excel add-in using Angular</span></span>

<span data-ttu-id="96768-103">Dans cet article, vous allez découvrir le processus de création d’un complément Excel à l’aide d’Angular et de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="96768-103">In this article, you'll walk you through the process of building an Excel add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="96768-104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96768-104">Prerequisites</span></span>

- [<span data-ttu-id="96768-105">Node.js</span><span class="sxs-lookup"><span data-stu-id="96768-105">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="96768-106">Installez la dernière version de [Yeoman](https://github.com/yeoman/yo) et le [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) globalement.</span><span class="sxs-lookup"><span data-stu-id="96768-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-web-app"></a><span data-ttu-id="96768-107">Création de l’application web</span><span class="sxs-lookup"><span data-stu-id="96768-107">Create the web app</span></span>

1. <span data-ttu-id="96768-108">Utilisez le générateur Yeoman pour créer un projet de complément Excel.</span><span class="sxs-lookup"><span data-stu-id="96768-108">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="96768-109">Exécutez la commande suivante, puis répondez aux invites comme suit :</span><span class="sxs-lookup"><span data-stu-id="96768-109">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="96768-110">**Sélectionnez un type de projet :** `Office Add-in project using Angular framework`</span><span class="sxs-lookup"><span data-stu-id="96768-110">**Choose a project type:** `Office Add-in project using Angular framework`</span></span>
    - <span data-ttu-id="96768-111">**Sélectionnez un type de script :** `Typescript`</span><span class="sxs-lookup"><span data-stu-id="96768-111">**Choose a script type:** `Typescript`</span></span>
    - <span data-ttu-id="96768-112">**Comment souhaitez-vous nommer votre complément ? :** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="96768-112">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="96768-113">**Quelle application client Office voulez-vous prendre en charge ? :**`Excel`</span><span class="sxs-lookup"><span data-stu-id="96768-113">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Générateur Yeoman](../images/yo-office-excel-angular.png)
    
    <span data-ttu-id="96768-115">Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants de nœud de la prise en charge.</span><span class="sxs-lookup"><span data-stu-id="96768-115">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

2. <span data-ttu-id="96768-116">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="96768-116">Navigate to the root folder of the project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="96768-117">Mise à jour du code</span><span class="sxs-lookup"><span data-stu-id="96768-117">Update the code</span></span>

1. <span data-ttu-id="96768-118">Dans votre éditeur de code, ouvrez le fichier **app.css**, ajoutez les styles suivants à la fin du fichier et enregistrez-le.</span><span class="sxs-lookup"><span data-stu-id="96768-118">In your code editor, open the file **app.css**, add the following styles to the end of the file, and save the file.</span></span>

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
        font-family: Arial;
        padding-top: 25px;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
        font-family: Arial;
    }

    .padding {
        padding: 15px;
    }

    .padding-sm {
        padding: 4px;
    }

    .normal-button {
        width: 80px;
        padding: 2px;
    }
    ```

2. <span data-ttu-id="96768-119">Ouvrez le fichier **src/app/app.component.html**, remplacez-en tout le contenu par le code suivant et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="96768-119">Open the file **src/app/app.component.html**, replace the entire contents with the following code, and save the file.</span></span>

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>{{welcomeMessage}}</h1>
        </div>
    </div>
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <br />
            <div role="button" class="ms-Button" (click)="setColor()">
                <span class="ms-Button-label">Set color</span>
                <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--ChevronRight"></i></span>
            </div>
        </div>
    </div>
    ```

3. <span data-ttu-id="96768-120">Ouvrez le fichier **src/app/app.component.ts**, remplacez-en tout le contenu par le code suivant et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="96768-120">Open the file **src/app/app.component.ts**, replace the entire contents with the following code, and save the file.</span></span>

    ```typescript
    import { Component } from '@angular/core';
    import * as OfficeHelpers from '@microsoft/office-js-helpers';

    const template = require('./app.component.html');

    @Component({
        selector: 'app-home',
        template
    })
    export default class AppComponent {
        welcomeMessage = 'Welcome';

        async setColor() {
            try {
                await Excel.run(async context => {
                    const range = context.workbook.getSelectedRange();
                    range.load('address');
                    range.format.fill.color = 'green';
                    await context.sync();
                    console.log(`The range address was ${range.address}.`);
                });
            } catch (error) {
                OfficeHelpers.UI.notify(error);
                OfficeHelpers.Utilities.log(error);
            }
        }

    }
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="96768-121">Mise à jour du manifeste</span><span class="sxs-lookup"><span data-stu-id="96768-121">Update the manifest</span></span>

1. <span data-ttu-id="96768-122">Ouvrez le fichier nommé **manifest.xml** pour définir les paramètres et les fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="96768-122">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span> 

2. <span data-ttu-id="96768-123">L’élément `ProviderName` possède une valeur d’espace réservé.</span><span class="sxs-lookup"><span data-stu-id="96768-123">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="96768-124">Remplacez-le par votre nom.</span><span class="sxs-lookup"><span data-stu-id="96768-124">Replace it with your name.</span></span>

3. <span data-ttu-id="96768-125">L’attribut `DefaultValue` de l’élément `Description` possède un espace réservé.</span><span class="sxs-lookup"><span data-stu-id="96768-125">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="96768-126">Remplacez-le par **A task pane add-in for Excel**.</span><span class="sxs-lookup"><span data-stu-id="96768-126">Replace it with **A task pane add-in for Excel**.</span></span>

4. <span data-ttu-id="96768-127">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="96768-127">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="96768-128">Démarrage du serveur de développement</span><span class="sxs-lookup"><span data-stu-id="96768-128">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a><span data-ttu-id="96768-129">Try it out</span><span class="sxs-lookup"><span data-stu-id="96768-129">Try it out</span></span>

1. <span data-ttu-id="96768-130">Suivez les instructions pour la plateforme que vous utiliserez pour exécuter votre complément et chargez une version test du complément dans Excel.</span><span class="sxs-lookup"><span data-stu-id="96768-130">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="96768-131">Windows : [Chargement de version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="96768-131">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="96768-132">Excel Online : [Chargement de versions test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="96768-132">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="96768-133">iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="96768-133">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

   
2. <span data-ttu-id="96768-134">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="96768-134">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Bouton Complément Excel](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="96768-136">Sélectionnez une plage de cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="96768-136">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="96768-137">Dans le volet Office, cliquez sur le bouton **Définir couleur** pour définir la couleur de la plage sélectionnée en vert.</span><span class="sxs-lookup"><span data-stu-id="96768-137">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Complément Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="96768-139">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="96768-139">Next steps</span></span>

<span data-ttu-id="96768-p104">Félicitations, vous avez créé un complément Excel à l’aide d’Angular ! Découvrez à présent les fonctionnalités des compléments Excel et créez un complément plus complexe en continuant le didacticiel sur le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="96768-p104">Congratulations, you've successfully created an Excel add-in using Angular! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="96768-142">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="96768-142">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="96768-143">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="96768-143">See also</span></span>

* [<span data-ttu-id="96768-144">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="96768-144">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="96768-145">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="96768-145">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="96768-146">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="96768-146">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="96768-147">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="96768-147">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)

