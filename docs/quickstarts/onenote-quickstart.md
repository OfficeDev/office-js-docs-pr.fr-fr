---
title: Créer votre premier complément du volet Office de OneNote
description: Découvrez comment créer un complément simple de volet des tâches OneNote simple à l’aide de l’API JavaScript pour Office.
ms.date: 07/07/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 6f057d650451d12e834d8f875f40d9d6d71ee4d7
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094154"
---
# <a name="build-your-first-onenote-task-pane-add-in"></a><span data-ttu-id="b3103-103">Créer votre premier complément du volet Office de OneNote</span><span class="sxs-lookup"><span data-stu-id="b3103-103">Build your first OneNote task pane add-in</span></span>

<span data-ttu-id="b3103-104">Cet article décrit comment créer un complément du volet Office de OneNote.</span><span class="sxs-lookup"><span data-stu-id="b3103-104">In this article, you'll walk through the process of building a OneNote task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="b3103-105">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="b3103-105">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="b3103-106">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="b3103-106">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="b3103-107">**Sélectionnez un type de projet :** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="b3103-107">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="b3103-108">**Sélectionnez un type de script :** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="b3103-108">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="b3103-109">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="b3103-109">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="b3103-110">**Quelle application client Office voulez-vous prendre en charge ?**</span><span class="sxs-lookup"><span data-stu-id="b3103-110">**Which Office client application would you like to support?**</span></span> `OneNote`

![Capture d’écran des invites et des réponses relatives au générateur Yeoman](../images/yo-office-onenote.png)

<span data-ttu-id="b3103-112">Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="b3103-112">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="b3103-113">Explorer le projet</span><span class="sxs-lookup"><span data-stu-id="b3103-113">Explore the project</span></span>

<span data-ttu-id="b3103-114">Le projet de complément que vous avez créé à l’aide du générateur Yeoman contient un exemple de code pour un complément de volet de tâches très simple.</span><span class="sxs-lookup"><span data-stu-id="b3103-114">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="b3103-115">Le fichier **./manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="b3103-115">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="b3103-116">Le fichier **./src/taskpane/taskpane.html** contient les balises HTML du volet Office.</span><span class="sxs-lookup"><span data-stu-id="b3103-116">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="b3103-117">Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet Office.</span><span class="sxs-lookup"><span data-stu-id="b3103-117">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="b3103-118">Le fichier **./src/taskpane/taskpane.js** contient le code de l’API JavaScript pour Office qui facilite l’interaction entre le volet Office et l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="b3103-118">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="b3103-119">Mise à jour du code</span><span class="sxs-lookup"><span data-stu-id="b3103-119">Update the code</span></span>

<span data-ttu-id="b3103-120">Ouvrez le fichier **./src/taskpane/taskpane.js** dans l’éditeur de code et ajoutez le code suivant à la fonction `run`.</span><span class="sxs-lookup"><span data-stu-id="b3103-120">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the `run` function.</span></span> <span data-ttu-id="b3103-121">Ce code utilise l’API JavaScript OneNote pour définir le titre de la page et ajouter un plan au corps de celle-ci.</span><span class="sxs-lookup"><span data-stu-id="b3103-121">This code uses the OneNote JavaScript API to set the page title and add an outline to the body of the page.</span></span>

```js
try {
    await OneNote.run(async context => {

        // Get the current page.
        var page = context.application.getActivePage();

        // Queue a command to set the page title.
        page.title = "Hello World";

        // Queue a command to add an outline to the page.
        var html = "<p><ol><li>Item #1</li><li>Item #2</li></ol></p>";
        page.addOutline(40, 90, html);

        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync();
    });
} catch (error) {
    console.log("Error: " + error);
}
```

## <a name="try-it-out"></a><span data-ttu-id="b3103-122">Essayez</span><span class="sxs-lookup"><span data-stu-id="b3103-122">Try it out</span></span>

1. <span data-ttu-id="b3103-123">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="b3103-123">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="b3103-124">Démarrez le serveur web local et chargez indépendamment votre complément.</span><span class="sxs-lookup"><span data-stu-id="b3103-124">Start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="b3103-125">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="b3103-125">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="b3103-126">Si vous êtes invité à installer un certificat après avoir exécuté une des commandes suivantes, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="b3103-126">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    > [!TIP]
    > <span data-ttu-id="b3103-127">Si vous testez votre complément sur Mac, exécutez la commande suivante avant de continuer.</span><span class="sxs-lookup"><span data-stu-id="b3103-127">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="b3103-128">Lorsque vous exécutez cette commande, le serveur web local démarre.</span><span class="sxs-lookup"><span data-stu-id="b3103-128">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    <span data-ttu-id="b3103-129">Exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="b3103-129">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="b3103-130">Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution).</span><span class="sxs-lookup"><span data-stu-id="b3103-130">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

3. <span data-ttu-id="b3103-131">Dans [OneNote sur le web](https://www.onenote.com/notebooks), ouvrez un bloc-notes, puis créez une page.</span><span class="sxs-lookup"><span data-stu-id="b3103-131">In [OneNote on the web](https://www.onenote.com/notebooks), open a notebook and create a new page.</span></span>

4. <span data-ttu-id="b3103-132">Choisissez **Insertion > Compléments Office** pour ouvrir la boîte de dialogue Compléments Office.</span><span class="sxs-lookup"><span data-stu-id="b3103-132">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="b3103-133">Si vous êtes connecté avec votre compte de consommateur, sélectionnez l’onglet **MES COMPLÉMENTS**, puis choisissez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="b3103-133">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="b3103-134">Si vous êtes connecté avec votre compte professionnel ou scolaire, sélectionnez l’onglet **MON ORGANISATION**, puis choisissez **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="b3103-134">If you're signed in with your work or education account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span> 

    <span data-ttu-id="b3103-135">L’image suivante montre l’onglet **MES COMPLÉMENTS** pour les blocs-notes de consommateurs.</span><span class="sxs-lookup"><span data-stu-id="b3103-135">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

5. <span data-ttu-id="b3103-136">Dans la boîte de dialogue Télécharger le complément, accédez à **manifest.xml** dans le dossier de projet, puis choisissez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="b3103-136">In the Upload Add-in dialog, browse to **manifest.xml** in your project folder, and then choose **Upload**.</span></span> 

6. <span data-ttu-id="b3103-137">Dans l’onglet **Accueil**, choisissez le bouton **Afficher le volet de tâches** du ruban.</span><span class="sxs-lookup"><span data-stu-id="b3103-137">From the **Home** tab, choose the **Show Taskpane** button in the ribbon.</span></span> <span data-ttu-id="b3103-138">Le volet Office du complément s’ouvre dans un iFrame à côté de la page OneNote.</span><span class="sxs-lookup"><span data-stu-id="b3103-138">The add-in task pane opens in an iFrame next to the OneNote page.</span></span>

7. <span data-ttu-id="b3103-139">Au bas du volet Office, sélectionnez le lien **Exécuter** pour définir le titre de la page et ajouter un plan au corps de celle-ci.</span><span class="sxs-lookup"><span data-stu-id="b3103-139">At the bottom of the task pane, choose the **Run** link to set the page title and add an outline to the body of the page.</span></span>

    ![Complément OneNote généré à partir de cette procédure pas à pas](../images/onenote-first-add-in-4.png)

## <a name="next-steps"></a><span data-ttu-id="b3103-141">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="b3103-141">Next steps</span></span>

<span data-ttu-id="b3103-142">Félicitations ! Vous avez créé un complément du volet Office de OneNote !</span><span class="sxs-lookup"><span data-stu-id="b3103-142">Congratulations, you've successfully created a OneNote task pane add-in!</span></span> <span data-ttu-id="b3103-143">Ensuite, vous allez étudier en détail les concepts fondamentaux de la création de compléments Excel.</span><span class="sxs-lookup"><span data-stu-id="b3103-143">Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="b3103-144">Vue d’ensemble de la programmation de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="b3103-144">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="b3103-145">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b3103-145">See also</span></span>

* [<span data-ttu-id="b3103-146">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="b3103-146">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="b3103-147">Création de compléments Office</span><span class="sxs-lookup"><span data-stu-id="b3103-147">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="b3103-148">Développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="b3103-148">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="b3103-149">Vue d’ensemble de la programmation de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="b3103-149">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="b3103-150">Référence de l’API JavaScript de OneNote</span><span class="sxs-lookup"><span data-stu-id="b3103-150">OneNote JavaScript API reference</span></span>](../reference/overview/onenote-add-ins-javascript-reference.md)
- [<span data-ttu-id="b3103-151">Exemple de grille d’évaluation</span><span class="sxs-lookup"><span data-stu-id="b3103-151">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)

