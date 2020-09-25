---
title: Création de votre premier complément Outlook
description: Découvrez comment créer un complément de volet des tâches Outlook simple à l’aide de l’API JavaScript pour Office.
ms.date: 09/22/2020
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: f4a3827b630ccee7cd8cef6222bfe6bac82f8ba2
ms.sourcegitcommit: fd110305c2be8660ab8a47c1da3e3969bd1ede86
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/23/2020
ms.locfileid: "48214609"
---
# <a name="build-your-first-outlook-add-in"></a><span data-ttu-id="d5c41-103">Création de votre premier complément Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c41-103">Build your first Outlook add-in</span></span>

<span data-ttu-id="d5c41-104">Dans cet article, vous découvrirez comment créer un complément du volet Office Outlook qui affiche au moins une propriété d’un message sélectionné.</span><span class="sxs-lookup"><span data-stu-id="d5c41-104">In this article, you'll walk through the process of building an Outlook task pane add-in that displays at least one property of a selected message.</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="d5c41-105">Créer le complément</span><span class="sxs-lookup"><span data-stu-id="d5c41-105">Create the add-in</span></span>

<span data-ttu-id="d5c41-106">Vous pouvez créer un complément Office à l’aide du [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) ou de Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="d5c41-106">You can create an Office Add-in by using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) or Visual Studio.</span></span> <span data-ttu-id="d5c41-107">Le générateur Yeoman crée un projet Node.js qui peut être géré avec du Visual Studio Code ou n’importe quel autre éditeur, alors que Visual Studio crée une solution Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="d5c41-107">The Yeoman generator creates a Node.js project that can be managed with Visual Studio Code or any other editor, whereas Visual Studio creates a Visual Studio solution.</span></span>  <span data-ttu-id="d5c41-108">Sélectionnez l’onglet correspondant à votre choix, puis suivez les instructions de création de votre complément et testez-le localement.</span><span class="sxs-lookup"><span data-stu-id="d5c41-108">Select the tab for the one you'd like to use and then follow the instructions to create your add-in and test it locally.</span></span>

# <a name="yeoman-generator"></a>[<span data-ttu-id="d5c41-109">Générateur Yeoman</span><span class="sxs-lookup"><span data-stu-id="d5c41-109">Yeoman generator</span></span>](#tab/yeomangenerator)

### <a name="prerequisites"></a><span data-ttu-id="d5c41-110">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="d5c41-110">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]

- <span data-ttu-id="d5c41-111">[Node.js](https://nodejs.org/) (la dernière version [LTS](https://nodejs.org/about/releases))</span><span class="sxs-lookup"><span data-stu-id="d5c41-111">[Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>

- <span data-ttu-id="d5c41-112">La dernière version de[Yeoman](https://github.com/yeoman/yo) et de [Yeoman Générateur de compléments Office](https://github.com/OfficeDev/generator-office). Pour installer ces outils globalement, exécutez la commande suivante à partir de l’invite de commande :</span><span class="sxs-lookup"><span data-stu-id="d5c41-112">The latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="d5c41-113">Même si vous avez précédemment installé le générateur Yeoman, nous vous recommandons de mettre à jour votre package vers la dernière version de npm.</span><span class="sxs-lookup"><span data-stu-id="d5c41-113">Even if you've previously installed the Yeoman generator, we recommend you update your package to the latest version from npm.</span></span>

### <a name="create-the-add-in-project"></a><span data-ttu-id="d5c41-114">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="d5c41-114">Create the add-in project</span></span>

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - <span data-ttu-id="d5c41-115">**Sélectionnez un type de projet** - `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="d5c41-115">**Choose a project type** - `Office Add-in Task Pane project`</span></span>

    - <span data-ttu-id="d5c41-116">**Sélectionnez un type de script** - `Javascript`</span><span class="sxs-lookup"><span data-stu-id="d5c41-116">**Choose a script type** - `Javascript`</span></span>

    - <span data-ttu-id="d5c41-117">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="d5c41-117">**What do you want to name your add-in?**</span></span> - `My Office Add-in`

    - <span data-ttu-id="d5c41-118">**Quelle application client Office voulez-vous prendre en charge ?**</span><span class="sxs-lookup"><span data-stu-id="d5c41-118">**Which Office client application would you like to support?**</span></span> - `Outlook`

    ![Capture d’écran des invites et des réponses relatives au générateur Yeoman](../images/yo-office-outlook.png)
    
    <span data-ttu-id="d5c41-120">Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="d5c41-120">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

1. <span data-ttu-id="d5c41-121">Accédez au dossier racine du projet de l’application web.</span><span class="sxs-lookup"><span data-stu-id="d5c41-121">Navigate to the root folder of the web application project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

### <a name="explore-the-project"></a><span data-ttu-id="d5c41-122">Explorer le projet</span><span class="sxs-lookup"><span data-stu-id="d5c41-122">Explore the project</span></span>

<span data-ttu-id="d5c41-123">Le projet de complément que vous avez créé à l’aide du générateur Yeoman contient un exemple de code pour un complément de volet de tâches très simple.</span><span class="sxs-lookup"><span data-stu-id="d5c41-123">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="d5c41-124">Le fichier **./manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="d5c41-124">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="d5c41-125">Le fichier **./src/taskpane/taskpane.html** contient les balises HTML du volet Office.</span><span class="sxs-lookup"><span data-stu-id="d5c41-125">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="d5c41-126">Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet Office.</span><span class="sxs-lookup"><span data-stu-id="d5c41-126">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="d5c41-127">Le fichier **./src/taskpane/taskpane.js** contient le code d’API JavaScript pour Office qui facilite l’interaction entre le volet Office et Outlook.</span><span class="sxs-lookup"><span data-stu-id="d5c41-127">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and Outlook.</span></span>

### <a name="update-the-code"></a><span data-ttu-id="d5c41-128">Mettre à jour le code</span><span class="sxs-lookup"><span data-stu-id="d5c41-128">Update the code</span></span>

1. <span data-ttu-id="d5c41-129">Dans votre éditeur de code, ouvrez le fichier **./src/taskpane/taskpane.html** et remplacez l’élément `<main>` (dans l’élément `<body>`) par le balisage suivant.</span><span class="sxs-lookup"><span data-stu-id="d5c41-129">In your code editor, open the file **./src/taskpane/taskpane.html** and replace the entire `<main>` element (within the `<body>` element) with the following markup.</span></span> <span data-ttu-id="d5c41-130">Ce nouveau balisage ajoute une étiquette à l’emplacement où le script dans **./src/taskpane/taskpane.js** écrira des données.</span><span class="sxs-lookup"><span data-stu-id="d5c41-130">This new markup adds a label where the script in **./src/taskpane/taskpane.js** will write data.</span></span>

    ```html
    <main id="app-body" class="ms-welcome__main" style="display: none;">
        <h2 class="ms-font-xl"> Discover what Office Add-ins can do for you today! </h2>
        <p><label id="item-subject"></label></p>
        <div role="button" id="run" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
            <span class="ms-Button-label">Run</span>
        </div>
    </main>
    ```

1. <span data-ttu-id="d5c41-131">Ouvrez le fichier **./src/taskpane/taskpane.js** dans l’éditeur de code et ajoutez le code suivant à la fonction `run`.</span><span class="sxs-lookup"><span data-stu-id="d5c41-131">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the `run` function.</span></span> <span data-ttu-id="d5c41-132">Ce code utilise l’API JavaScript pour Office afin d’obtenir une référence au message en cours et écrire sa valeur de propriété `subject` dans le volet Office.</span><span class="sxs-lookup"><span data-stu-id="d5c41-132">This code uses the Office JavaScript API to get a reference to the current message and write its `subject` property value to the task pane.</span></span>

    ```js
    // Get a reference to the current message
    var item = Office.context.mailbox.item;

    // Write message property value to the task pane
    document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
    ```

### <a name="try-it-out"></a><span data-ttu-id="d5c41-133">Essayez !</span><span class="sxs-lookup"><span data-stu-id="d5c41-133">Try it out</span></span>

> [!NOTE]
> <span data-ttu-id="d5c41-134">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="d5c41-134">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="d5c41-135">Si vous êtes invité à installer un certificat après avoir exécuté la commande suivante, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="d5c41-135">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span> <span data-ttu-id="d5c41-136">Il se peut également que vous deviez exécuter votre invite de commande ou votre terminal en tant qu'administrateur pour que les modifications soient effectuées.</span><span class="sxs-lookup"><span data-stu-id="d5c41-136">You may also have to run your command prompt or terminal as an administrator for the changes to be made.</span></span>

1. <span data-ttu-id="d5c41-137">Exécutez la commande suivante dans le répertoire racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="d5c41-137">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="d5c41-138">Lorsque vous exécutez cette commande, le serveur web local démarre (s’il n’est pas déjà en cours d’exécution).</span><span class="sxs-lookup"><span data-stu-id="d5c41-138">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="d5c41-139">Suivez les instructions indiquées dans l’article [Chargement de version test des compléments Outlook](../outlook/sideload-outlook-add-ins-for-testing.md) pour charger le complément dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="d5c41-139">Follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md) to sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="d5c41-140">Sur Outlook, affichez un message dans le [volet de lecture](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0)ou ouvrez le message dans sa propre fenêtre.</span><span class="sxs-lookup"><span data-stu-id="d5c41-140">In Outlook, view a message in the [Reading Pane](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0), or open the message in its own window.</span></span>

1. <span data-ttu-id="d5c41-141">Sélectionnez l’onglet **Accueil** (ou l’onglet **Message** si vous avez ouvert le message dans une nouvelle fenêtre), puis sélectionnez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="d5c41-141">Choose the **Home** tab (or the **Message** tab if you opened the message in a new window), and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Capture d’écran d’une fenêtre de message dans Outlook avec le bouton du complément mis en surbrillance](../images/quick-start-button-1.png)

    > [!NOTE]
    > <span data-ttu-id="d5c41-143">Si le message d’erreur « Désolé... nous ne pouvons pas ouvrir ce complément à partir de localhost » s’affiche dans le volet Office, suivez les étapes décrites dans l’[article résolution des problèmes](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost).</span><span class="sxs-lookup"><span data-stu-id="d5c41-143">If you receive the error "We can't open this add-in from localhost" in the task pane, follow the steps outlined in the [troubleshooting article](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost).</span></span>

1. <span data-ttu-id="d5c41-144">Faites défiler vers le bas du volet Office et sélectionnez le lien **Exécuter** pour écrire l’objet du message dans le volet Office.</span><span class="sxs-lookup"><span data-stu-id="d5c41-144">Scroll to the bottom of the task pane and choose the **Run** link to write the message subject to the task pane.</span></span>

    ![Capture d’écran du volet Office du complément avec le lien d’exécution mis en évidence](../images/quick-start-task-pane-2.png)

    ![Capture d’écran du volet Office du complément, affichant le sujet du message](../images/quick-start-task-pane-3.png)

### <a name="next-steps"></a><span data-ttu-id="d5c41-147">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="d5c41-147">Next steps</span></span>

<span data-ttu-id="d5c41-148">Félicitations, vous avez créé votre premier complément de volet de tâches Outlook !</span><span class="sxs-lookup"><span data-stu-id="d5c41-148">Congratulations, you've successfully created your first Outlook task pane add-in!</span></span> <span data-ttu-id="d5c41-149">Ensuite, découvrez les fonctionnalités d’un complément Outlook et créez-en un plus complexe en suivant le [didacticiel pour complément Outlook](../tutorials/outlook-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="d5c41-149">Next, learn more about the capabilities of an Outlook add-in and build a more complex add-in by following along with the [Outlook add-in tutorial](../tutorials/outlook-tutorial.md).</span></span>

# <a name="visual-studio"></a>[<span data-ttu-id="d5c41-150">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="d5c41-150">Visual Studio</span></span>](#tab/visualstudio)

### <a name="prerequisites"></a><span data-ttu-id="d5c41-151">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="d5c41-151">Prerequisites</span></span>

- <span data-ttu-id="d5c41-152">[Visual Studio 2019](https://www.visualstudio.com/vs/) avec la charge de travail de **développement Office/SharePoint** installée</span><span class="sxs-lookup"><span data-stu-id="d5c41-152">[Visual Studio 2019](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!NOTE]
    > <span data-ttu-id="d5c41-153">Si vous avez déjà installé Visual Studio 2019, [utilisez Visual Studio Installer](/visualstudio/install/modify-visual-studio) pour vérifier que la charge de travail de **développement Office/SharePoint** est bien installée.</span><span class="sxs-lookup"><span data-stu-id="d5c41-153">If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span>

- <span data-ttu-id="d5c41-154">Office 365</span><span class="sxs-lookup"><span data-stu-id="d5c41-154">Office 365</span></span>

    > [!NOTE]
    > <span data-ttu-id="d5c41-155">Si vous n’avez pas d’abonnement Microsoft 365, vous pouvez en obtenir un gratuitement en vous inscrivant au [programme développeur Microsoft 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="d5c41-155">If you do not have a Microsoft 365 subscription, you can get a free one by signing up for the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

### <a name="create-the-add-in-project"></a><span data-ttu-id="d5c41-156">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="d5c41-156">Create the add-in project</span></span>

1. <span data-ttu-id="d5c41-157">Dans la barre de menu de Visual Studio, choisissez successivement **Fichier** > **Nouveau** > **Projet**.</span><span class="sxs-lookup"><span data-stu-id="d5c41-157">On the Visual Studio menu bar, choose **File** > **New** > **Project**.</span></span>

1. <span data-ttu-id="d5c41-158">Dans la liste des types de projets sous **Visual C#** ou **Visual Basic**, développez **Office/SharePoint**, choisissez **Compléments**, puis **Complément web Outlook** pour le type de projet.</span><span class="sxs-lookup"><span data-stu-id="d5c41-158">In the list of project types under **Visual C#** or **Visual Basic**, expand **Office/SharePoint**, choose **Add-ins**, and then choose **Outlook Web Add-in** as the project type.</span></span>

1. <span data-ttu-id="d5c41-159">Nommez le projet, puis cliquez sur **OK**.</span><span class="sxs-lookup"><span data-stu-id="d5c41-159">Name the project, and then choose **OK**.</span></span>

1. <span data-ttu-id="d5c41-160">Visual Studio crée une solution et ses deux projets apparaissent dans l’**Explorateur de solutions**.</span><span class="sxs-lookup"><span data-stu-id="d5c41-160">Visual Studio creates a solution and its two projects appear in **Solution Explorer**.</span></span> <span data-ttu-id="d5c41-161">Le fichier **MessageRead.html** s’ouvre dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="d5c41-161">The **MessageRead.html** file opens in Visual Studio.</span></span>

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="d5c41-162">Explorer la solution Visual Studio</span><span class="sxs-lookup"><span data-stu-id="d5c41-162">Explore the Visual Studio solution</span></span>

<span data-ttu-id="d5c41-163">Quand vous arrivez au bout de l’Assistant, Visual Studio crée une solution qui contient deux projets.</span><span class="sxs-lookup"><span data-stu-id="d5c41-163">When you've completed the wizard, Visual Studio creates a solution that contains two projects.</span></span>

|<span data-ttu-id="d5c41-164">**Project**</span><span class="sxs-lookup"><span data-stu-id="d5c41-164">**Project**</span></span>|<span data-ttu-id="d5c41-165">**Description**</span><span class="sxs-lookup"><span data-stu-id="d5c41-165">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="d5c41-166">Projet de complément</span><span class="sxs-lookup"><span data-stu-id="d5c41-166">Add-in project</span></span>|<span data-ttu-id="d5c41-167">Contient uniquement un fichier manifeste XML contenant tous les paramètres qui décrivent votre complément.</span><span class="sxs-lookup"><span data-stu-id="d5c41-167">Contains only an XML manifest file, which contains all the settings that describe your add-in.</span></span> <span data-ttu-id="d5c41-168">Ces paramètres aident l’application Office à déterminer le moment où votre complément doit être activé et l’emplacement où il doit apparaître.</span><span class="sxs-lookup"><span data-stu-id="d5c41-168">These settings help the Office application determine when your add-in should be activated and where the add-in should appear.</span></span> <span data-ttu-id="d5c41-169">Visual Studio génère le contenu de ce fichier pour vous permettre d’exécuter le projet et d’utiliser votre complément immédiatement.</span><span class="sxs-lookup"><span data-stu-id="d5c41-169">Visual Studio generates the contents of this file for you so that you can run the project and use your add-in immediately.</span></span> <span data-ttu-id="d5c41-170">Vous pouvez modifier ces paramètres à tout moment en modifiant le fichier XML.</span><span class="sxs-lookup"><span data-stu-id="d5c41-170">You can change these settings any time by modifying the XML file.</span></span>|
|<span data-ttu-id="d5c41-171">Projet d’application web</span><span class="sxs-lookup"><span data-stu-id="d5c41-171">Web application project</span></span>|<span data-ttu-id="d5c41-p109">Contient les pages de contenu de votre complément, notamment tous les fichiers et références de fichiers dont vous avez besoin pour développer des pages HTML et JavaScript compatibles avec Office. Pendant que vous développez votre complément, Visual Studio héberge l’application web sur votre serveur IIS local. Lorsque vous êtes prêt à publier le complément, vous devez déployer ce projet d’application web sur un serveur web.</span><span class="sxs-lookup"><span data-stu-id="d5c41-p109">Contains the content pages of your add-in, including all the files and file references that you need to develop Office-aware HTML and JavaScript pages. While you develop your add-in, Visual Studio hosts the web application on your local IIS server. When you're ready to publish the add-in, you'll need to deploy this web application project to a web server.</span></span>|

### <a name="update-the-code"></a><span data-ttu-id="d5c41-175">Mise à jour du code</span><span class="sxs-lookup"><span data-stu-id="d5c41-175">Update the code</span></span>

1. <span data-ttu-id="d5c41-176">**MessageRead.html** spécifie le code HTML qui s’affichera dans le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="d5c41-176">**MessageRead.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="d5c41-177">Dans **MessageRead.html**, remplacez l’élément `<body>` par les marques suivantes et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="d5c41-177">In **MessageRead.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
    ```HTML
    <body class="ms-font-m ms-welcome">
        <div class="ms-Fabric content-main">
            <h1 class="ms-font-xxl">Message properties</h1>
            <table class="ms-Table ms-Table--selectable">
                <thead>
                    <tr>
                        <th>Property</th>
                        <th>Value</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td><strong>Id</strong></td>
                        <td class="prop-val"><code><label id="item-id"></label></code></td>
                    </tr>
                    <tr>
                        <td><strong>Subject</strong></td>
                        <td class="prop-val"><code><label id="item-subject"></label></code></td>
                    </tr>
                    <tr>
                        <td><strong>Message Id</strong></td>
                        <td class="prop-val"><code><label id="item-internetMessageId"></label></code></td>
                    </tr>
                    <tr>
                        <td><strong>From</strong></td>
                        <td class="prop-val"><code><label id="item-from"></label></code></td>
                    </tr>
                </tbody>
            </table>
        </div>
    </body>
    ```

1. <span data-ttu-id="d5c41-178">Ouvrez le fichier **MessageRead.js** à la racine du projet d’application web.</span><span class="sxs-lookup"><span data-stu-id="d5c41-178">Open the file **MessageRead.js** in the root of the web application project.</span></span> <span data-ttu-id="d5c41-179">Ce fichier spécifie le script pour le complément.</span><span class="sxs-lookup"><span data-stu-id="d5c41-179">This file specifies the script for the add-in.</span></span> <span data-ttu-id="d5c41-180">Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="d5c41-180">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {

        Office.onReady(function () {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                loadItemProps(Office.context.mailbox.item);
            });
        });

        function loadItemProps(item) {
            // Write message property values to the task pane
            $('#item-id').text(item.itemId);
            $('#item-subject').text(item.subject);
            $('#item-internetMessageId').text(item.internetMessageId);
            $('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");
        }
    })();
    ```

1. <span data-ttu-id="d5c41-181">Ouvrez le fichier **MessageRead.css** à la racine du projet d’application web.</span><span class="sxs-lookup"><span data-stu-id="d5c41-181">Open the file **MessageRead.css** in the root of the web application project.</span></span> <span data-ttu-id="d5c41-182">Ce fichier spécifie les styles personnalisés pour le complément.</span><span class="sxs-lookup"><span data-stu-id="d5c41-182">This file specifies the custom styles for the add-in.</span></span> <span data-ttu-id="d5c41-183">Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="d5c41-183">Replace the entire contents with the following code and save the file.</span></span>

    ```CSS
    html,
    body {
        width: 100%;
        height: 100%;
        margin: 0;
        padding: 0;
    }

    td.prop-val {
        word-break: break-all;
    }

    .content-main {
        margin: 10px;
    }
    ```

### <a name="update-the-manifest"></a><span data-ttu-id="d5c41-184">Mise à jour du manifeste</span><span class="sxs-lookup"><span data-stu-id="d5c41-184">Update the manifest</span></span>

1. <span data-ttu-id="d5c41-p113">Ouvrez le fichier manifeste XML dans le projet de complément. Ce fichier définit les paramètres et les fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="d5c41-p113">Open the XML manifest file in the Add-in project. This file defines the add-in's settings and capabilities.</span></span>

1. <span data-ttu-id="d5c41-p114">L’élément `ProviderName` possède une valeur d’espace réservé. Remplacez-le par votre nom.</span><span class="sxs-lookup"><span data-stu-id="d5c41-p114">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

1. <span data-ttu-id="d5c41-189">L’attribut `DefaultValue` de l’élément `DisplayName` possède un espace réservé.</span><span class="sxs-lookup"><span data-stu-id="d5c41-189">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="d5c41-190">Remplacez-le par `My Office Add-in`.</span><span class="sxs-lookup"><span data-stu-id="d5c41-190">Replace it with `My Office Add-in`.</span></span>

1. <span data-ttu-id="d5c41-191">L’attribut `DefaultValue` de l’élément `Description` possède un espace réservé.</span><span class="sxs-lookup"><span data-stu-id="d5c41-191">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="d5c41-192">Remplacez-le par `My First Outlook add-in`.</span><span class="sxs-lookup"><span data-stu-id="d5c41-192">Replace it with `My First Outlook add-in`.</span></span>

1. <span data-ttu-id="d5c41-193">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="d5c41-193">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="My First Outlook add-in"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="d5c41-194">Try it out</span><span class="sxs-lookup"><span data-stu-id="d5c41-194">Try it out</span></span>

1. <span data-ttu-id="d5c41-195">À l’aide de Visual Studio, testez le complément Outlook que vous venez de créer en appuyant sur F5 ou en sélectionnant le bouton **Démarrer**.</span><span class="sxs-lookup"><span data-stu-id="d5c41-195">Using Visual Studio, test the newly created Outlook add-in by pressing F5 or choosing the **Start** button.</span></span> <span data-ttu-id="d5c41-196">Le complément est hébergé localement sur IIS.</span><span class="sxs-lookup"><span data-stu-id="d5c41-196">The add-in will be hosted locally on IIS.</span></span>

1. <span data-ttu-id="d5c41-197">Dans la boîte de dialogue**Se connecter à un compte de messagerie Exchange**, entrez l’adresse de messagerie et mot de passe pour votre [compte Microsoft](https://account.microsoft.com/account), puis sélectionnez**Se connecter**.</span><span class="sxs-lookup"><span data-stu-id="d5c41-197">In the **Connect to Exchange email account** dialog box, enter the email address and password for your [Microsoft account](https://account.microsoft.com/account) and then choose **Connect**.</span></span> <span data-ttu-id="d5c41-198">Lorsque la page de connexion Outlook.com s’ouvre dans un navigateur, connectez-vous à votre compte de courrier avec les mêmes informations d’identification que vous avez entrées précédemment.</span><span class="sxs-lookup"><span data-stu-id="d5c41-198">When the Outlook.com login page opens in a browser, sign in to your email account with the same credentials as you entered previously.</span></span>

    > [!NOTE]
    > <span data-ttu-id="d5c41-199">Si la boîte de dialogue **Se connecter au compte de messagerie Exchange** vous invite à vous connecter à plusieurs reprises, l’authentification de base est peut-être désactivée pour les comptes sur votre client Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="d5c41-199">If the **Connect to Exchange email account** dialog box repeatedly prompts you to sign in, Basic Auth may be disabled for accounts on your Microsoft 365 tenant.</span></span> <span data-ttu-id="d5c41-200">Pour tester ce complément, connectez-vous à l’aide d’un [compte Microsoft](https://account.microsoft.com/account) à la place.</span><span class="sxs-lookup"><span data-stu-id="d5c41-200">To test this add-in, sign in using a [Microsoft account](https://account.microsoft.com/account) instead.</span></span>

1. <span data-ttu-id="d5c41-201">Dans Outlook sur le web, sélectionnez ou ouvrez un message.</span><span class="sxs-lookup"><span data-stu-id="d5c41-201">In Outlook on the web, select or open a message.</span></span>

1. <span data-ttu-id="d5c41-202">Dans le message, recherchez les points de suspension du menu de dépassement de capacité contenant le bouton du complément.</span><span class="sxs-lookup"><span data-stu-id="d5c41-202">Within the message, locate the ellipsis for the overflow menu containing the add-in's button.</span></span>

    ![Capture d’écran d’une fenêtre de message dans Outlook sur le web avec les points de suspension mis en surbrillance](../images/quick-start-button-owa-1.png)

1. <span data-ttu-id="d5c41-204">Dans le menu de dépassement de capacité, recherchez le bouton du complément.</span><span class="sxs-lookup"><span data-stu-id="d5c41-204">Within the overflow menu, locate the add-in's button.</span></span>

    ![Capture d’écran d’une fenêtre de message dans Outlook sur le web avec le bouton du complément mis en surbrillance](../images/quick-start-button-owa-2.png)

1. <span data-ttu-id="d5c41-206">Cliquez sur le bouton pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="d5c41-206">Click the button to open the add-in's task pane.</span></span>

    ![Capture d’écran du volet Office du complément dans Outlook sur le web, affichant les propriétés des messages](../images/quick-start-task-pane-owa-1.png)

    > [!NOTE]
    > <span data-ttu-id="d5c41-208">Si le volet Office n’est pas chargé, essayez de l’ouvrir dans un navigateur sur le même ordinateur.</span><span class="sxs-lookup"><span data-stu-id="d5c41-208">If the task pane doesn't load, try to verify by opening it in a browser on the same machine.</span></span>

### <a name="next-steps"></a><span data-ttu-id="d5c41-209">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="d5c41-209">Next steps</span></span>

<span data-ttu-id="d5c41-210">Félicitations, vous avez créé votre premier complément de volet de tâches Outlook !</span><span class="sxs-lookup"><span data-stu-id="d5c41-210">Congratulations, you've successfully created your first Outlook task pane add-in!</span></span> <span data-ttu-id="d5c41-211">Ensuite, en savoir plus sur la [création de compléments Office avec Visual Studio](../develop/develop-add-ins-visual-studio.md).</span><span class="sxs-lookup"><span data-stu-id="d5c41-211">Next, learn more about [developing Office Add-ins with Visual Studio](../develop/develop-add-ins-visual-studio.md).</span></span>

---
