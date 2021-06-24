---
title: Créer votre premier complément du volet des tâches de Project
description: Découvrez comment créer un complément simple de volet des tâches Project à l’aide de l’API JavaScript pour Office.
ms.date: 06/07/2021
ms.prod: project
localization_priority: Priority
ms.openlocfilehash: 3589572c022e4fe3dc75cd25e5d378a3dab0c8b7
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076573"
---
# <a name="build-your-first-project-task-pane-add-in"></a><span data-ttu-id="167eb-103">Créer votre premier complément du volet des tâches de Project</span><span class="sxs-lookup"><span data-stu-id="167eb-103">Build your first Project task pane add-in</span></span>

<span data-ttu-id="167eb-104">Cet article décrit comment créer un complément du volet des tâches de Project.</span><span class="sxs-lookup"><span data-stu-id="167eb-104">In this article, you'll walk through the process of building a Project task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="167eb-105">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="167eb-105">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="167eb-106">Project 2016 ou version ultérieure pour Windows</span><span class="sxs-lookup"><span data-stu-id="167eb-106">Project 2016 or later on Windows</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="167eb-107">Créer le complément</span><span class="sxs-lookup"><span data-stu-id="167eb-107">Create the add-in</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="167eb-108">**Sélectionnez un type de projet :** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="167eb-108">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="167eb-109">**Sélectionnez un type de script :** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="167eb-109">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="167eb-110">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="167eb-110">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="167eb-111">**Quelle application client Office voulez-vous prendre en charge ?**</span><span class="sxs-lookup"><span data-stu-id="167eb-111">**Which Office client application would you like to support?**</span></span> `Project`

![Capture d'écran montrant les invites et les réponses pour le générateur Yeoman dans une interface de ligne de commande.](../images/yo-office-project.png)

<span data-ttu-id="167eb-113">Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="167eb-113">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="167eb-114">Explorer le projet</span><span class="sxs-lookup"><span data-stu-id="167eb-114">Explore the project</span></span>

<span data-ttu-id="167eb-115">Le projet de complément que vous avez créé à l’aide du générateur Yeoman contient un exemple de code pour un complément de volet de tâches très simple.</span><span class="sxs-lookup"><span data-stu-id="167eb-115">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span>

- <span data-ttu-id="167eb-116">Le fichier **./manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="167eb-116">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="167eb-117">Le fichier **./src/taskpane/taskpane.html** contient les balises HTML du volet Office.</span><span class="sxs-lookup"><span data-stu-id="167eb-117">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="167eb-118">Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet Office.</span><span class="sxs-lookup"><span data-stu-id="167eb-118">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="167eb-119">Le fichier **./src/taskpane/taskpane.js** contient le code d’API JavaScript pour Office qui facilite l’interaction entre le volet des tâches et l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="167eb-119">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="167eb-120">Mettre à jour le code</span><span class="sxs-lookup"><span data-stu-id="167eb-120">Update the code</span></span>

<span data-ttu-id="167eb-121">Ouvrez le fichier **./src/taskpane/taskpane.js** dans votre éditeur de code et ajoutez le code suivant à la fonction `run`.</span><span class="sxs-lookup"><span data-stu-id="167eb-121">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the `run` function.</span></span> <span data-ttu-id="167eb-122">Ce code utilise l’API JavaScript Office pour définir le champ `Name` et le champ `Notes` de la tâche sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="167eb-122">This code uses the Office JavaScript API to set the `Name` field and `Notes` field of the selected task.</span></span>

```js
var taskGuid;

// Get the GUID of the selected task
Office.context.document.getSelectedTaskAsync(
    function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            taskGuid = result.value;

            // Set the specified fields for the selected task.
            var targetFields = [Office.ProjectTaskFields.Name, Office.ProjectTaskFields.Notes];
            var fieldValues = ['New task name', 'Notes for the task.'];

            // Set the field value. If the call is successful, set the next field.
            for (var i = 0; i < targetFields.length; i++) {
                Office.context.document.setTaskFieldAsync(
                    taskGuid,
                    targetFields[i],
                    fieldValues[i],
                    function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            i++;
                        }
                        else {
                            var err = result.error;
                            console.log(err.name + ' ' + err.code + ' ' + err.message);
                        }
                    }
                );
            }
        } else {
            var err = result.error;
            console.log(err.name + ' ' + err.code + ' ' + err.message);
        }
    }
);
```

## <a name="try-it-out"></a><span data-ttu-id="167eb-123">Essayez</span><span class="sxs-lookup"><span data-stu-id="167eb-123">Try it out</span></span>

1. <span data-ttu-id="167eb-124">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="167eb-124">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="167eb-125">Démarrez le serveur web local.</span><span class="sxs-lookup"><span data-stu-id="167eb-125">Start the local web server.</span></span>

    > [!NOTE]
    > <span data-ttu-id="167eb-126">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="167eb-126">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="167eb-127">Si vous êtes invité à installer un certificat après avoir exécuté la commande suivante, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="167eb-127">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    <span data-ttu-id="167eb-128">Exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="167eb-128">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="167eb-129">Lorsque vous exécutez cette commande, le serveur web local démarre.</span><span class="sxs-lookup"><span data-stu-id="167eb-129">When you run this command, the local web server will start.</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

3. <span data-ttu-id="167eb-130">Dans Project, créez un plan de projet simple.</span><span class="sxs-lookup"><span data-stu-id="167eb-130">In Project, create a simple project plan.</span></span>

4. <span data-ttu-id="167eb-131">Chargez votre complément dans Project en suivant les instructions fournies dans [Chargement de versions test de compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="167eb-131">Load your add-in in Project by following the instructions in [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

5. <span data-ttu-id="167eb-132">Sélectionnez une seule tâche dans le projet.</span><span class="sxs-lookup"><span data-stu-id="167eb-132">Select a single task within the project.</span></span>

6. <span data-ttu-id="167eb-133">Au bas du volet des tâches, sélectionnez le lien **Exécuter** pour renommer la tâche sélectionnée et ajouter des notes à la tâche sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="167eb-133">At the bottom of the task pane, choose the **Run** link to rename the selected task and add notes to the selected task.</span></span>

    ![Capture d'écran de l'application Project avec le complément du volet des tâches chargé.](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a><span data-ttu-id="167eb-135">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="167eb-135">Next steps</span></span>

<span data-ttu-id="167eb-p104">Félicitations, vous avez créé un complément de volet de tâches Project ! À présent, découvrez les fonctionnalités d’un complément Project et explorez des scénarios courants.</span><span class="sxs-lookup"><span data-stu-id="167eb-p104">Congratulations, you've successfully created a Project task pane add-in! Next, learn more about the capabilities of a Project add-in and explore common scenarios.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="167eb-138">Compléments Project</span><span class="sxs-lookup"><span data-stu-id="167eb-138">Project add-ins</span></span>](../project/project-add-ins.md)

## <a name="see-also"></a><span data-ttu-id="167eb-139">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="167eb-139">See also</span></span>

- [<span data-ttu-id="167eb-140">Développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="167eb-140">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="167eb-141">Concepts de base pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="167eb-141">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
