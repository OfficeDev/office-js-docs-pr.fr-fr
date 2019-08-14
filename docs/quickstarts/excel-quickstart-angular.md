---
title: Créer un complément de volet de tâches Excel à l’aide d’Angular
description: ''
ms.date: 05/02/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 5d27b7d70fb013ca15115af3ad550a604ce10bb9
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/13/2019
ms.locfileid: "36308035"
---
# <a name="build-an-excel-task-pane-add-in-using-angular"></a><span data-ttu-id="fe61c-102">Créer un complément de volet de tâches Excel à l’aide d’Angular</span><span class="sxs-lookup"><span data-stu-id="fe61c-102">Build an Excel task pane add-in using Angular</span></span>

<span data-ttu-id="fe61c-103">Cet article décrit le processus de création d’un complément de volet de tâches Excel à l’aide d’Angular et de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="fe61c-103">In this article, you'll walk through the process of building an Excel task pane add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="fe61c-104">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="fe61c-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="fe61c-105">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="fe61c-105">Create the add-in project</span></span>

[!include[note about Yeoman generator bug](../includes/note-yeoman-generator-bug-201908.md)]

<span data-ttu-id="fe61c-106">Utilisez le générateur Yeoman pour créer un projet de complément Excel.</span><span class="sxs-lookup"><span data-stu-id="fe61c-106">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="fe61c-107">Exécutez la commande suivante, puis répondez aux invites comme suit :</span><span class="sxs-lookup"><span data-stu-id="fe61c-107">Run the following command and then answer the prompts as follows:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="fe61c-108">**Sélectionnez un type de projet :** `Office Add-in Task Pane project using Angular framework`</span><span class="sxs-lookup"><span data-stu-id="fe61c-108">**Choose a project type:** `Office Add-in Task Pane project using Angular framework`</span></span>
- <span data-ttu-id="fe61c-109">**Sélectionnez un type de script :** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="fe61c-109">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="fe61c-110">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="fe61c-110">**What do you want to name your add-in?**</span></span> `my-office-add-in`
- <span data-ttu-id="fe61c-111">**Quelle application client Office voulez-vous prendre en charge ?**</span><span class="sxs-lookup"><span data-stu-id="fe61c-111">**Which Office client application would you like to support?**</span></span> `Excel`

<span data-ttu-id="fe61c-112">Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="fe61c-112">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

## <a name="explore-the-project"></a><span data-ttu-id="fe61c-113">Explorer le projet</span><span class="sxs-lookup"><span data-stu-id="fe61c-113">Explore the project</span></span>

<span data-ttu-id="fe61c-114">Le projet de complément que vous avez créé à l’aide du générateur Yeoman contient un exemple de code pour un complément de volet de tâches très simple.</span><span class="sxs-lookup"><span data-stu-id="fe61c-114">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> <span data-ttu-id="fe61c-115">Pour explorer les composants clés de votre projet de complément, ouvrez le projet dans votre éditeur de code et passez en revue les fichiers répertoriés ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="fe61c-115">If you'd like to explore the key components of your add-in project, open the project in your code editor and review the files listed below.</span></span> <span data-ttu-id="fe61c-116">Lorsque vous êtes prêt à tester votre complément, passez à la section suivante.</span><span class="sxs-lookup"><span data-stu-id="fe61c-116">When you're ready to try out your add-in, proceed to the next section.</span></span>

- <span data-ttu-id="fe61c-117">Le fichier **manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="fe61c-117">The **manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="fe61c-118">Le fichier **./src/taskpane/app/app.component.html** contient les balises HTML du volet Office.</span><span class="sxs-lookup"><span data-stu-id="fe61c-118">The **./src/taskpane/app/app.component.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="fe61c-119">Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet Office.</span><span class="sxs-lookup"><span data-stu-id="fe61c-119">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="fe61c-120">Le fichier **./src/taskpane/app/app.component.ts** contient le code d’API JavaScript pour Office qui facilite l’interaction entre le volet Office et Excel.</span><span class="sxs-lookup"><span data-stu-id="fe61c-120">The **./src/taskpane/app/app.component.ts** file contains the Office JavaScript API code that facilitates interaction between the task pane and Excel.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="fe61c-121">Essayez</span><span class="sxs-lookup"><span data-stu-id="fe61c-121">Try it out</span></span>

1. <span data-ttu-id="fe61c-122">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="fe61c-122">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "my-office-add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. <span data-ttu-id="fe61c-123">Dans Excel, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="fe61c-123">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Bouton Complément Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="fe61c-125">Sélectionnez une plage de cellules dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="fe61c-125">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="fe61c-126">En bas du volet Office, cliquez sélectionnez le lien **Exécuter** pour définir la couleur de la plage sélectionnée sur jaune.</span><span class="sxs-lookup"><span data-stu-id="fe61c-126">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Complément Excel](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a><span data-ttu-id="fe61c-128">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="fe61c-128">Next steps</span></span>

<span data-ttu-id="fe61c-129">Félicitations, vous avez créé un complément de volet de tâches Excel à l’aide d’Angular !</span><span class="sxs-lookup"><span data-stu-id="fe61c-129">Congratulations, you've successfully created an Excel task pane add-in using Angular!</span></span> <span data-ttu-id="fe61c-130">Ensuite, découvrez les fonctionnalités d’un complément Excel et créez-en un plus complexe en suivant le didacticiel sur les compléments Excel.</span><span class="sxs-lookup"><span data-stu-id="fe61c-130">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="fe61c-131">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="fe61c-131">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="fe61c-132">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="fe61c-132">See also</span></span>

* [<span data-ttu-id="fe61c-133">Didacticiel sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="fe61c-133">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="fe61c-134">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="fe61c-134">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="fe61c-135">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="fe61c-135">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="fe61c-136">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="fe61c-136">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
