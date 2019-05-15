---
title: Créer un complément de volet de tâches Excel à l’aide d’Angular
description: ''
ms.date: 05/02/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 66c85ba9914b783295e9ed2143dc9ce107f64c4c
ms.sourcegitcommit: 47b792755e655043d3db2f1fdb9a1eeb7453c636
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33619907"
---
# <a name="build-an-excel-task-pane-add-in-using-angular"></a><span data-ttu-id="e96ad-102">Créer un complément de volet de tâches Excel à l’aide d’Angular</span><span class="sxs-lookup"><span data-stu-id="e96ad-102">Build an Excel add-in using Angular</span></span>

<span data-ttu-id="e96ad-103">Cet article décrit le processus de création d’un complément de volet de tâches Excel à l’aide d’Angular et de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="e96ad-103">In this article, you'll walk through the process of building an Excel add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="e96ad-104">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="e96ad-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="e96ad-105">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="e96ad-105">Create the add-in project</span></span>

1. <span data-ttu-id="e96ad-106">Utilisez le générateur Yeoman pour créer un projet de complément Excel.</span><span class="sxs-lookup"><span data-stu-id="e96ad-106">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="e96ad-107">Exécutez la commande suivante, puis répondez aux invites comme suit :</span><span class="sxs-lookup"><span data-stu-id="e96ad-107">Run the following command and then answer the prompts as follows:</span></span>

    ```command&nbsp;line
    yo office
    ```

    - <span data-ttu-id="e96ad-108">**Sélectionnez un type de projet :** `Office Add-in Task Pane project using Angular framework`</span><span class="sxs-lookup"><span data-stu-id="e96ad-108">**Choose a project type:** `Office Add-in Task Pane project using Angular framework`</span></span>
    - <span data-ttu-id="e96ad-109">**Sélectionnez un type de script :** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="e96ad-109">**Choose a script type:** `TypeScript`</span></span>
    - <span data-ttu-id="e96ad-110">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="e96ad-110">**What do you want to name your add-in?**</span></span> `My Office Add-in`
    - <span data-ttu-id="e96ad-111">**Quelle application client Office voulez-vous prendre en charge ?**</span><span class="sxs-lookup"><span data-stu-id="e96ad-111">**Which Office client application would you like to support?**</span></span> `Excel`

    ![Générateur Yeoman](../images/yo-office-excel-angular-2.png)

    <span data-ttu-id="e96ad-113">Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants de nœud de la prise en charge.</span><span class="sxs-lookup"><span data-stu-id="e96ad-113">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

2. <span data-ttu-id="e96ad-114">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="e96ad-114">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```
## <a name="explore-the-project"></a><span data-ttu-id="e96ad-115">Explorer le projet</span><span class="sxs-lookup"><span data-stu-id="e96ad-115">Explore the project</span></span>

<span data-ttu-id="e96ad-116">Le projet de complément que vous avez créé à l’aide du générateur Yeoman contient un exemple de code pour un complément de volet de tâches très simple.</span><span class="sxs-lookup"><span data-stu-id="e96ad-116">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> <span data-ttu-id="e96ad-117">Pour explorer les composants clés de votre projet de complément, ouvrez le projet dans votre éditeur de code et passez en revue les fichiers répertoriés ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="e96ad-117">If you'd like to explore the key components of your add-in project, open the project in your code editor and review the files listed below.</span></span> <span data-ttu-id="e96ad-118">Lorsque vous êtes prêt à tester votre complément, passez à la section suivante.</span><span class="sxs-lookup"><span data-stu-id="e96ad-118">When you're ready to try out your add-in, proceed to the next section.</span></span>

- <span data-ttu-id="e96ad-119">Le fichier **manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="e96ad-119">The **manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="e96ad-120">Le fichier **./src/taskpane/app/app.component.html** contient les balises HTML du volet Office.</span><span class="sxs-lookup"><span data-stu-id="e96ad-120">The **./src/taskpane/app/app.component.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="e96ad-121">Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet Office.</span><span class="sxs-lookup"><span data-stu-id="e96ad-121">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="e96ad-122">Le fichier **./src/taskpane/app/app.component.ts** contient le code d’API JavaScript pour Office qui facilite l’interaction entre le volet Office et Excel.</span><span class="sxs-lookup"><span data-stu-id="e96ad-122">The **./src/taskpane/app/app.component.ts** file contains the Office JavaScript API code that facilitates interaction between the task pane and Excel.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="e96ad-123">Essayez !</span><span class="sxs-lookup"><span data-stu-id="e96ad-123">Try it out</span></span>

1. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

2. <span data-ttu-id="e96ad-124">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="e96ad-124">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Bouton Complément Excel](../images/excel-quickstart-addin-3b.png)

3. <span data-ttu-id="e96ad-126">Sélectionnez une plage de cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="e96ad-126">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="e96ad-127">En bas du volet Office, cliquez sélectionnez le lien **Exécuter** pour définir la couleur de la plage sélectionnée sur jaune.</span><span class="sxs-lookup"><span data-stu-id="e96ad-127">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Complément Excel](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a><span data-ttu-id="e96ad-129">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="e96ad-129">Next steps</span></span>

<span data-ttu-id="e96ad-130">Félicitations, vous avez créé un complément de volet de tâches Excel à l’aide d’Angular !</span><span class="sxs-lookup"><span data-stu-id="e96ad-130">Congratulations, you've successfully created an Excel add-in using Angular!</span></span> <span data-ttu-id="e96ad-131">Ensuite, découvrez les fonctionnalités d’un complément Excel et créez-en un plus complexe en suivant le didacticiel sur les compléments Excel.</span><span class="sxs-lookup"><span data-stu-id="e96ad-131">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="e96ad-132">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="e96ad-132">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="e96ad-133">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e96ad-133">See also</span></span>

* [<span data-ttu-id="e96ad-134">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="e96ad-134">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="e96ad-135">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="e96ad-135">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="e96ad-136">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="e96ad-136">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="e96ad-137">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="e96ad-137">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
