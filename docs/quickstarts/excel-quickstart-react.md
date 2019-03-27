---
title: Développement d’un complément Excel à l’aide de React
description: ''
ms.date: 03/19/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 6b56054337d14b8836e9b9994cbdfda8feee5b33
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871954"
---
# <a name="build-an-excel-add-in-using-react"></a><span data-ttu-id="ff1f5-102">Développement d’un complément Excel à l’aide de React</span><span class="sxs-lookup"><span data-stu-id="ff1f5-102">Build an Excel add-in using React</span></span>

<span data-ttu-id="ff1f5-103">Cet article décrit le processus de création d’un complément Excel à l’aide de React et de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-103">In this article, you'll walk through the process of building an Excel add-in using React and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="ff1f5-104">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="ff1f5-104">Prerequisites</span></span>

- [<span data-ttu-id="ff1f5-105">Node.js</span><span class="sxs-lookup"><span data-stu-id="ff1f5-105">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="ff1f5-106">Installez la dernière version de [Yeoman](https://github.com/yeoman/yo) et le [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) globalement.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>
    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-web-app"></a><span data-ttu-id="ff1f5-107">Création de l’application web</span><span class="sxs-lookup"><span data-stu-id="ff1f5-107">Create the web app</span></span>

1. <span data-ttu-id="ff1f5-108">Utilisez le générateur Yeoman pour créer un projet de complément Excel.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-108">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="ff1f5-109">Exécutez la commande suivante, puis répondez aux invites comme suit :</span><span class="sxs-lookup"><span data-stu-id="ff1f5-109">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="ff1f5-110">**Sélectionnez un type de projet :** `Office Add-in project using React framework`</span><span class="sxs-lookup"><span data-stu-id="ff1f5-110">**Choose a project type:** `Office Add-in project using React framework`</span></span>
    - <span data-ttu-id="ff1f5-111">**Comment souhaitez-vous nommer votre complément ? :** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="ff1f5-111">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="ff1f5-112">**Quelle application client Office voulez-vous prendre en charge ? :**`Excel`</span><span class="sxs-lookup"><span data-stu-id="ff1f5-112">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Générateur Yeoman](../images/yo-office-excel-react.png)

    <span data-ttu-id="ff1f5-114">Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants de nœud de la prise en charge.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-114">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

2. <span data-ttu-id="ff1f5-115">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-115">Navigate to the root folder of the project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="ff1f5-116">Mise à jour du code</span><span class="sxs-lookup"><span data-stu-id="ff1f5-116">Update the code</span></span>

1. <span data-ttu-id="ff1f5-117">Dans votre éditeur de code, ouvrez le fichier **src/styles.less**, ajoutez les styles suivants à la fin du fichier et enregistrez-le.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-117">In your code editor, open the file **src/styles.less**, add the following styles to the end of the file, and save the file.</span></span>

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

2. <span data-ttu-id="ff1f5-118">Le modèle de projet créés dans le générateur Yeoman de compléments Office inclut un composant React qui n’est pas nécessaire pour ce guide de démarrage rapide.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-118">The project template that the Office Add-ins Yeoman generator created includes a React component that is not needed for this quick start.</span></span> <span data-ttu-id="ff1f5-119">Supprimez le fichier **src/components/HeroList.tsx**.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-119">Delete the file **src/components/HeroList.tsx**.</span></span>

3. <span data-ttu-id="ff1f5-120">Ouvrez le fichier **src/components/Header.tsx**, remplacez-en tout le contenu par le code suivant et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-120">Open the file **src/components/Header.tsx**, replace the entire contents with the following code, and save the file.</span></span>

    ```typescript
    import * as React from 'react';

    export interface HeaderProps {
        title: string;
    }

    export class Header extends React.Component<HeaderProps, any> {
        constructor(props, context) {
            super(props, context);
        }

        render() {
            return (
                <div id='content-header'>
                    <div className='padding'>
                        <h1>{this.props.title}</h1>
                    </div>
                </div>
            );
        }
    }
    ```

4. <span data-ttu-id="ff1f5-121">Créez un nouveau composant React nommé **Content.tsx** dans le dossier **src/composants**, ajoutez le code suivant et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-121">Create a new React component named **Content.tsx** in the **src/components** folder, add the following code, and save the file.</span></span>

    ```typescript
    import * as React from 'react';
    import { Button, ButtonType } from 'office-ui-fabric-react';

    export interface ContentProps {
        message: string;
        buttonLabel: string;
        click: any;
    }

    export class Content extends React.Component<ContentProps, any> {
        constructor(props, context) {
            super(props, context);
        }

        render() {
            return (
                <div id='content-main'>
                    <div className='padding'>
                        <p>{this.props.message}</p>
                        <br />
                        <h3>Try it out</h3>
                        <br/>
                        <Button className='normal-button' buttonType={ButtonType.hero} onClick={this.props.click}>{this.props.buttonLabel}</Button>
                    </div>
                </div>
            );
        }
    }
    ```

5. <span data-ttu-id="ff1f5-122">Ouvrez le fichier **src/components/App.tsx**, remplacez-en tout le contenu par le code suivant et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-122">Open the file **src/components/App.tsx**, replace the entire contents with the following code, and save the file.</span></span>

    ```typescript
    /* global Office, Excel */

    import * as React from 'react';
    import { Header } from './Header';
    import { Content } from './Content';
    import Progress from './Progress';

    import * as OfficeHelpers from '@microsoft/office-js-helpers';

    export interface AppProps {
        title: string;
        isOfficeInitialized: boolean;
    }

    export interface AppState {
    }

    export default class App extends React.Component<AppProps, AppState> {
        constructor(props, context) {
            super(props, context);
        }

        setColor = async () => {
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

        render() {
            const {
                title,
                isOfficeInitialized,
            } = this.props;

            if (!isOfficeInitialized) {
                return (
                    <Progress
                        title={title}
                        logo='assets/logo-filled.png'
                        message='Please sideload your add-in to see app body.'
                    />
                );
            }

            return (
                <div className='ms-welcome'>
                    <Header title='Welcome' />
                    <Content message='Choose the button below to set the color of the selected range to green.' buttonLabel='Set color' click={this.setColor} />
                </div>
            );
        }
    }
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="ff1f5-123">Mise à jour du manifeste</span><span class="sxs-lookup"><span data-stu-id="ff1f5-123">Update the manifest</span></span>

1. <span data-ttu-id="ff1f5-124">Ouvrez le fichier nommé **manifest.xml** pour définir les paramètres et les fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-124">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span> 

2. <span data-ttu-id="ff1f5-p103">L’élément `ProviderName` possède une valeur d’espace réservé. Remplacez-le par votre nom.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-p103">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="ff1f5-p104">L’attribut `DefaultValue` de l’élément `Description` possède un espace réservé. Remplacez-le par **A task pane add-in for Excel**.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-p104">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for Excel**.</span></span>

4. <span data-ttu-id="ff1f5-129">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-129">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="ff1f5-130">Démarrage du serveur de développement</span><span class="sxs-lookup"><span data-stu-id="ff1f5-130">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)]

## <a name="try-it-out"></a><span data-ttu-id="ff1f5-131">Essayez !</span><span class="sxs-lookup"><span data-stu-id="ff1f5-131">Try it out</span></span>

1. <span data-ttu-id="ff1f5-132">Suivez les instructions pour la plateforme que vous utiliserez afin d’exécuter votre complément en vue d’en charger une version test dans Excel.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-132">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="ff1f5-133">Windows : [Chargement de version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="ff1f5-133">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="ff1f5-134">Excel Online : [Chargement de versions test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="ff1f5-134">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="ff1f5-135">iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="ff1f5-135">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="ff1f5-136">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-136">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Bouton Complément Excel](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="ff1f5-138">Sélectionnez une plage de cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-138">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="ff1f5-139">Dans le volet Office, cliquez sur le bouton **Définir couleur** pour définir la couleur de la plage sélectionnée en vert.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-139">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Complément Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="ff1f5-141">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="ff1f5-141">Next steps</span></span>

<span data-ttu-id="ff1f5-p105">Félicitations, vous avez créé un complément Excel à l’aide de React ! Découvrez à présent les fonctionnalités des compléments Excel et créez un complément plus complexe en continuant le didacticiel sur le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="ff1f5-p105">Congratulations, you've successfully created an Excel add-in using React! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="ff1f5-144">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="ff1f5-144">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="ff1f5-145">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ff1f5-145">See also</span></span>

* [<span data-ttu-id="ff1f5-146">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="ff1f5-146">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="ff1f5-147">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="ff1f5-147">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="ff1f5-148">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="ff1f5-148">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="ff1f5-149">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="ff1f5-149">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
