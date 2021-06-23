---
title: Créer un complément de volet de tâches Excel à l’aide d’Angular
description: Découvrez comment créer un complément de volet des tâches Excel simple à l’aide de l’API JavaScript et d’Angular pour Office.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: d843a74b3542df8dbc462ae2876179de7b42a2d2
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076930"
---
# <a name="build-an-excel-task-pane-add-in-using-angular"></a><span data-ttu-id="01dc2-103">Créer un complément de volet de tâches Excel à l’aide d’Angular</span><span class="sxs-lookup"><span data-stu-id="01dc2-103">Build an Excel task pane add-in using Angular</span></span>

<span data-ttu-id="01dc2-104">Cet article décrit le processus de création d’un complément de volet de tâches Excel à l’aide d’Angular et de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="01dc2-104">In this article, you'll walk through the process of building an Excel task pane add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="01dc2-105">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="01dc2-105">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="01dc2-106">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="01dc2-106">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="01dc2-107">**Sélectionnez un type de projet :** `Office Add-in Task Pane project using Angular framework`</span><span class="sxs-lookup"><span data-stu-id="01dc2-107">**Choose a project type:** `Office Add-in Task Pane project using Angular framework`</span></span>
- <span data-ttu-id="01dc2-108">**Sélectionnez un type de script :** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="01dc2-108">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="01dc2-109">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="01dc2-109">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="01dc2-110">**Quelle application client Office voulez-vous prendre en charge ?**</span><span class="sxs-lookup"><span data-stu-id="01dc2-110">**Which Office client application would you like to support?**</span></span> `Excel`

![Capture d’écran de l’interface de ligne de commande du générateur de compléments Yeoman Office, avec l’option type de projet réglée sur l’infrastructure Angular.](../images/yo-office-excel-angular-2.png)

<span data-ttu-id="01dc2-112">Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="01dc2-112">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="01dc2-113">Explorer le projet</span><span class="sxs-lookup"><span data-stu-id="01dc2-113">Explore the project</span></span>

<span data-ttu-id="01dc2-114">Le projet de complément que vous avez créé à l’aide du générateur Yeoman contient un exemple de code pour un complément de volet de tâches très simple.</span><span class="sxs-lookup"><span data-stu-id="01dc2-114">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> <span data-ttu-id="01dc2-115">Pour explorer les composants clés de votre projet de complément, ouvrez le projet dans votre éditeur de code et passez en revue les fichiers répertoriés ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="01dc2-115">If you'd like to explore the key components of your add-in project, open the project in your code editor and review the files listed below.</span></span> <span data-ttu-id="01dc2-116">Lorsque vous êtes prêt à tester votre complément, passez à la section suivante.</span><span class="sxs-lookup"><span data-stu-id="01dc2-116">When you're ready to try out your add-in, proceed to the next section.</span></span>

- <span data-ttu-id="01dc2-117">Le fichier **manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="01dc2-117">The **manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="01dc2-118">Le fichier **./src/taskpane/app/app.component.html** contient les balises HTML du volet Office.</span><span class="sxs-lookup"><span data-stu-id="01dc2-118">The **./src/taskpane/app/app.component.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="01dc2-119">Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet Office.</span><span class="sxs-lookup"><span data-stu-id="01dc2-119">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="01dc2-120">Le fichier **./src/taskpane/app/app.component.ts** contient le code d’API JavaScript pour Office qui facilite l’interaction entre le volet Office et Excel.</span><span class="sxs-lookup"><span data-stu-id="01dc2-120">The **./src/taskpane/app/app.component.ts** file contains the Office JavaScript API code that facilitates interaction between the task pane and Excel.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="01dc2-121">Essayez</span><span class="sxs-lookup"><span data-stu-id="01dc2-121">Try it out</span></span>

1. <span data-ttu-id="01dc2-122">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="01dc2-122">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. <span data-ttu-id="01dc2-123">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="01dc2-123">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Capture d’écran du menu Accueil d’Excel, avec le bouton Afficher le volet Office mis en évidence.](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="01dc2-125">Sélectionnez une plage de cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="01dc2-125">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="01dc2-126">En bas du volet Office, cliquez sélectionnez le lien **Exécuter** pour définir la couleur de la plage sélectionnée sur jaune.</span><span class="sxs-lookup"><span data-stu-id="01dc2-126">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Capture d’écran d’Excel, avec le volet Office du complément ouvert et le bouton Exécuter mis en surbrillance dans ce volet.](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a><span data-ttu-id="01dc2-128">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="01dc2-128">Next steps</span></span>

<span data-ttu-id="01dc2-p102">Félicitations, vous avez créé un complément du volet Office Excel à l’aide d’Angular ! Maintenant, apprenez-en davantage sur les fonctionnalités d’un complément Excel et créez un complément plus complexe en suivant le didacticiel sur les compléments Excel.</span><span class="sxs-lookup"><span data-stu-id="01dc2-p102">Congratulations, you've successfully created an Excel task pane add-in using Angular! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="01dc2-131">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="01dc2-131">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="01dc2-132">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="01dc2-132">See also</span></span>

* [<span data-ttu-id="01dc2-133">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="01dc2-133">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="01dc2-134">Développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="01dc2-134">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="01dc2-135">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="01dc2-135">Excel JavaScript object model in Office Add-ins</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="01dc2-136">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="01dc2-136">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="01dc2-137">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="01dc2-137">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
