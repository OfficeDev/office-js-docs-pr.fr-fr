---
title: Créer votre premier complément du volet des tâches de Project
description: ''
ms.date: 09/06/2019
ms.prod: project
localization_priority: Priority
ms.openlocfilehash: 0a7684f3d1bd4f404ba42a798908bb9d2ba2f8d2
ms.sourcegitcommit: ce7e7087a4550b9c090dc565fee5eac08a2985a2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/06/2019
ms.locfileid: "36782281"
---
# <a name="build-your-first-project-task-pane-add-in"></a><span data-ttu-id="414dc-102">Créer votre premier complément du volet des tâches de Project</span><span class="sxs-lookup"><span data-stu-id="414dc-102">Build your first Project task pane add-in</span></span>

<span data-ttu-id="414dc-103">Cet article décrit comment créer un complément du volet des tâches de Project.</span><span class="sxs-lookup"><span data-stu-id="414dc-103">In this article, you'll walk through the process of building a Project task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="414dc-104">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="414dc-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="414dc-105">Project 2016 ou version ultérieure pour Windows</span><span class="sxs-lookup"><span data-stu-id="414dc-105">Project 2016 or later on Windows</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="414dc-106">Créer le complément</span><span class="sxs-lookup"><span data-stu-id="414dc-106">Create the add-in</span></span>

<span data-ttu-id="414dc-107">Utilisez le générateur Yeoman afin de créer un projet de complément Project.</span><span class="sxs-lookup"><span data-stu-id="414dc-107">Use the Yeoman generator to create a Project add-in project.</span></span> <span data-ttu-id="414dc-108">Exécutez la commande suivante, puis répondez aux invites comme suit :</span><span class="sxs-lookup"><span data-stu-id="414dc-108">Run the following command and then answer the prompts as follows:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="414dc-109">**Sélectionnez un type de projet :** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="414dc-109">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="414dc-110">**Sélectionnez un type de script :** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="414dc-110">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="414dc-111">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="414dc-111">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="414dc-112">**Quelle application client Office voulez-vous prendre en charge ?**</span><span class="sxs-lookup"><span data-stu-id="414dc-112">**Which Office client application would you like to support?**</span></span> `Project`

![Capture d’écran des invites et des réponses relatives au générateur Yeoman](../images/yo-office-project.png)

<span data-ttu-id="414dc-114">Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="414dc-114">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

## <a name="explore-the-project"></a><span data-ttu-id="414dc-115">Explorer le projet</span><span class="sxs-lookup"><span data-stu-id="414dc-115">Explore the project</span></span>

<span data-ttu-id="414dc-116">Le projet de complément que vous avez créé à l’aide du générateur Yeoman contient un exemple de code pour un complément de volet de tâches très simple.</span><span class="sxs-lookup"><span data-stu-id="414dc-116">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="414dc-117">Le fichier **./manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="414dc-117">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="414dc-118">Le fichier **./src/taskpane/taskpane.html** contient les balises HTML du volet Office.</span><span class="sxs-lookup"><span data-stu-id="414dc-118">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="414dc-119">Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet Office.</span><span class="sxs-lookup"><span data-stu-id="414dc-119">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="414dc-120">Le fichier **./src/taskpane/taskpane.js** contient le code de l’API JavaScript pour Office qui facilite l’interaction entre le volet Office et l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="414dc-120">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="414dc-121">Mettre à jour le code</span><span class="sxs-lookup"><span data-stu-id="414dc-121">Update the code</span></span>

<span data-ttu-id="414dc-122">Dans votre éditeur de code, ouvrez le fichier **./src/taskpane/taskpane.js** et ajoutez le code suivant à la fonction **run**.</span><span class="sxs-lookup"><span data-stu-id="414dc-122">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the **run** function.</span></span> <span data-ttu-id="414dc-123">Ce code utilise l’API JavaScript Office pour définir le champ `Name` et le champ `Notes` de la tâche sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="414dc-123">This code uses the Office JavaScript API to set the `Name` field and `Notes` field of the selected task.</span></span>

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

## <a name="try-it-out"></a><span data-ttu-id="414dc-124">Essayez</span><span class="sxs-lookup"><span data-stu-id="414dc-124">Try it out</span></span>

1. <span data-ttu-id="414dc-125">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="414dc-125">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="414dc-126">Démarrez le serveur web local.</span><span class="sxs-lookup"><span data-stu-id="414dc-126">Start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="414dc-127">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="414dc-127">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="414dc-128">Si vous êtes invité à installer un certificat après avoir exécuté la commande suivante, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="414dc-128">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    <span data-ttu-id="414dc-129">Exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="414dc-129">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="414dc-130">Lorsque vous exécutez cette commande, le serveur web local démarre.</span><span class="sxs-lookup"><span data-stu-id="414dc-130">When you run this command, the local web server will start.</span></span>

    ```command&nbsp;line
    npm start
    ```

3. <span data-ttu-id="414dc-131">Dans Project, créez un plan de projet simple.</span><span class="sxs-lookup"><span data-stu-id="414dc-131">In Project, create a simple project plan.</span></span>

4. <span data-ttu-id="414dc-132">Chargez votre complément dans Project en suivant les instructions fournies dans [Chargement de versions test de compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="414dc-132">Load your add-in in Project by following the instructions in [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

5. <span data-ttu-id="414dc-133">Sélectionnez une seule tâche dans le projet.</span><span class="sxs-lookup"><span data-stu-id="414dc-133">Select a single task within the project.</span></span>

6. <span data-ttu-id="414dc-134">Au bas du volet des tâches, sélectionnez le lien **Exécuter** pour renommer la tâche sélectionnée et ajouter des notes à la tâche sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="414dc-134">At the bottom of the task pane, choose the **Run** link to rename the selected task and add notes to the selected task.</span></span>

    ![Capture d’écran de l’application Project avec le complément du volet des tâches chargé](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a><span data-ttu-id="414dc-136">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="414dc-136">Next steps</span></span>

<span data-ttu-id="414dc-137">Félicitations, vous avez créé un complément du volet des tâches de Project !</span><span class="sxs-lookup"><span data-stu-id="414dc-137">Congratulations, you've successfully created a Project task pane add-in!</span></span> <span data-ttu-id="414dc-138">Ensuite, découvrez les fonctionnalités d’un complément Project et explorez des scénarios plus courants.</span><span class="sxs-lookup"><span data-stu-id="414dc-138">Next, learn more about the capabilities of a Project add-in and explore common scenarios.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="414dc-139">Compléments Project</span><span class="sxs-lookup"><span data-stu-id="414dc-139">Project add-ins</span></span>](../project/project-add-ins.md)

