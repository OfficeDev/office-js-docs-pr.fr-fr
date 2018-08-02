# <a name="build-your-first-project-add-in"></a><span data-ttu-id="e8869-101">Création de votre premier complément Project</span><span class="sxs-lookup"><span data-stu-id="e8869-101">Build your first Project add-in</span></span>

<span data-ttu-id="e8869-102">Cet article décrit le processus de création d’un complément Project à l’aide de jQuery et de l’API JavaScript pour Office.</span><span class="sxs-lookup"><span data-stu-id="e8869-102">In this article, you'll walk through the process of building a Project add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="e8869-103">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="e8869-103">Prerequisites</span></span>

- [<span data-ttu-id="e8869-104">Node.js</span><span class="sxs-lookup"><span data-stu-id="e8869-104">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="e8869-105">Installez la dernière version de [Yeoman](https://github.com/yeoman/yo) et le [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) globalement.</span><span class="sxs-lookup"><span data-stu-id="e8869-105">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in"></a><span data-ttu-id="e8869-106">Créer le complément</span><span class="sxs-lookup"><span data-stu-id="e8869-106">Create the add-in</span></span>

1. <span data-ttu-id="e8869-107">Créez un dossier sur votre lecteur local et nommez-le `my-project-addin`.</span><span class="sxs-lookup"><span data-stu-id="e8869-107">Create a folder on your local drive and name it `my-project-addin`.</span></span> <span data-ttu-id="e8869-108">Il s’agit de l’emplacement dans lequel vous allez créer les fichiers de votre complément.</span><span class="sxs-lookup"><span data-stu-id="e8869-108">This is where you'll create the files for your add-in.</span></span>

2. <span data-ttu-id="e8869-109">Accédez à votre nouveau dossier.</span><span class="sxs-lookup"><span data-stu-id="e8869-109">Navigate to your new folder.</span></span>

    ```bash
    cd my-project-addin
    ```

3. <span data-ttu-id="e8869-110">Utilisez le générateur Yeoman afin de créer un projet de complément Project.</span><span class="sxs-lookup"><span data-stu-id="e8869-110">Use the Yeoman generator to create a Project add-in project.</span></span> <span data-ttu-id="e8869-111">Exécutez la commande suivante, puis répondez aux invites comme suit :</span><span class="sxs-lookup"><span data-stu-id="e8869-111">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="e8869-112">**Choisissez un type de projet :** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="e8869-112">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="e8869-113">**Choisissez un type de script :** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="e8869-113">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="e8869-114">**Comment souhaitez-vous nommer votre complément ? :** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="e8869-114">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="e8869-115">**Quelle application client Office voulez-vous prendre en charge ? :** `Project`</span><span class="sxs-lookup"><span data-stu-id="e8869-115">**Which Office client application would you like to support?:** `Project`</span></span>

    ![Capture d’écran des invites et des réponses relatives au générateur Yeoman](../images/yo-office-project-jquery.png)
    
    <span data-ttu-id="e8869-117">Une fois que vous avez terminé avec l'assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="e8869-117">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="e8869-118">Mise à jour du code</span><span class="sxs-lookup"><span data-stu-id="e8869-118">Update the code</span></span>

1. <span data-ttu-id="e8869-119">Dans votre éditeur de code, ouvrez **index.html** à la racine du projet.</span><span class="sxs-lookup"><span data-stu-id="e8869-119">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="e8869-120">Ce fichier contient le code HTML qui s’affichera dans le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="e8869-120">This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="e8869-121">Remplacez l’élément `<header>` à l’intérieur de l’élément `<body>` par le balisage suivant.</span><span class="sxs-lookup"><span data-stu-id="e8869-121">Replace the `<header>` element inside the `<body>` element with the following markup.</span></span>

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    ```

3. <span data-ttu-id="e8869-122">Remplacez l’élément `<main>` dans l’élément `<body>` par le balisage suivant et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="e8869-122">Replace the `<main>` element inside the `<body>` element with the following markup and save the file.</span></span>

    ```html
    <div id="content-main">
        <div class="padding">
            <p>Select a task and then choose the buttons below and observe the output in the <b>Results</b> textbox.</p>
            <h3>Try it out</h3>
            <button class="ms-Button" id="get-task-guid">Get Task GUID</button>
            <br/><br/>
            <button class="ms-Button" id="get-task">Get Task data</button>
            <br/>
            <h4>Results:</h4>
            <textarea id="result" rows="6" cols="25"></textarea>
        </div>
    </div>
    ```

4. <span data-ttu-id="e8869-123">Ouvrez le fichier **app.js** pour spécifier le script pour le complément.</span><span class="sxs-lookup"><span data-stu-id="e8869-123">Open the file **app.js** to specify the script for the add-in.</span></span> <span data-ttu-id="e8869-124">Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="e8869-124">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {

        var taskGuid;

        // The initialize function must be run each time a new page is loaded
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#get-task-guid').click(getTaskGUID);
                $('#get-task').click(getTask);
            });
        };

        function getTaskGUID() {
            Office.context.document.getSelectedTaskAsync(function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    result.value = "Task GUID: " + asyncResult.value;
                    taskGuid = asyncResult.value;
                }
                else {
                    console.log(asyncResult.error.message);
                }
            });
        }

        function getTask() {
            if (taskGuid != undefined) {
                Office.context.document.getTaskAsync(
                    taskGuid,
                    function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            var taskInfo = asyncResult.value;
                            var taskOutput = "Task name: " + taskInfo.taskName +
                                            "\nGUID: " + taskGuid +
                                            "\nWSS Id: " + taskInfo.wssTaskId +
                                            "\nResource names: " + taskInfo.resourceNames;
                            result.value = taskOutput;
                        } else {
                            console.log(asyncResult.error.message);
                        }
                    }
                );
            } else {
                result.value = 'Task GUID not valid:\n' + taskGuid;
            } 
        }
    })();
    ```

4. <span data-ttu-id="e8869-125">Ouvrez le fichier **app.css** à la racine du projet pour spécifier les styles personnalisés du complément.</span><span class="sxs-lookup"><span data-stu-id="e8869-125">Open the file **app.css** in the root of the project to specify the custom styles for the add-in.</span></span> <span data-ttu-id="e8869-126">Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="e8869-126">Replace the entire contents with the following and save the file.</span></span>

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

## <a name="update-the-manifest"></a><span data-ttu-id="e8869-127">Mise à jour du manifeste</span><span class="sxs-lookup"><span data-stu-id="e8869-127">Update the manifest</span></span>

1. <span data-ttu-id="e8869-128">Ouvrez le fichier nommé **my-office-add-in-manifest.xml** pour définir les paramètres et les fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="e8869-128">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="e8869-129">L’élément `ProviderName` possède une valeur d’espace réservé.</span><span class="sxs-lookup"><span data-stu-id="e8869-129">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="e8869-130">Remplacez-le par votre nom.</span><span class="sxs-lookup"><span data-stu-id="e8869-130">Replace it with your name.</span></span>

3. <span data-ttu-id="e8869-131">L’attribut `DefaultValue` de l’élément `Description` possède un espace réservé.</span><span class="sxs-lookup"><span data-stu-id="e8869-131">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="e8869-132">Remplacez-le par **A task pane add-in for Project**.</span><span class="sxs-lookup"><span data-stu-id="e8869-132">Replace it with **A task pane add-in for Project**.</span></span>

4. <span data-ttu-id="e8869-133">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="e8869-133">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Project"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="e8869-134">Démarrage du serveur de développement</span><span class="sxs-lookup"><span data-stu-id="e8869-134">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a><span data-ttu-id="e8869-135">Essayez !</span><span class="sxs-lookup"><span data-stu-id="e8869-135">Try it out</span></span>

1. <span data-ttu-id="e8869-136">Dans Project, créez un projet simple comportant au moins une tâche.</span><span class="sxs-lookup"><span data-stu-id="e8869-136">In Project, create a simple project that has at least one task.</span></span>

2. <span data-ttu-id="e8869-137">Suivez les instructions pour la plateforme que vous utiliserez afin d’exécuter votre complément en vue d’en charger une version test dans Project.</span><span class="sxs-lookup"><span data-stu-id="e8869-137">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Project.</span></span>

    - <span data-ttu-id="e8869-138">Windows : [Chargement de version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="e8869-138">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="e8869-139">Project Online : [Chargement de version test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="e8869-139">Project Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="e8869-140">iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="e8869-140">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

3. <span data-ttu-id="e8869-141">Dans Project, sélectionnez une tâche.</span><span class="sxs-lookup"><span data-stu-id="e8869-141">In Project, select a task.</span></span>

    ![Capture d’écran d’un plan de projet dans Project avec une tâche sélectionnée](../images/project_quickstart_addin_1.png)

4. <span data-ttu-id="e8869-143">Dans le volet Office, sélectionnez le bouton **Get Task GUID** pour écrire le GUID de la tâche dans la zone de texte **Results**.</span><span class="sxs-lookup"><span data-stu-id="e8869-143">In the task pane, choose the **Get Task GUID** button to write the task GUID to the **Results** textbox.</span></span>

    ![Capture d’écran d’un plan de projet dans Project avec une tâche sélectionnée et le GUID de la tâche écrit dans la zone de texte dans le volet Office](../images/project_quickstart_addin_2.png)

5. <span data-ttu-id="e8869-145">Dans le volet Office, sélectionnez le bouton **Get Task data** pour écrire plusieurs propriétés de la tâche sélectionnée dans la zone de texte **Results**.</span><span class="sxs-lookup"><span data-stu-id="e8869-145">In the task pane, choose the **Get Task data** button to write several properties of the selected task to the **Results** textbox.</span></span>

    ![Capture d’écran d’un plan de projet dans Project avec une tâche sélectionnée et plusieurs propriétés de la tâche écrites dans la zone de texte dans le volet Office](../images/project_quickstart_addin_3.png)

## <a name="next-steps"></a><span data-ttu-id="e8869-147">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="e8869-147">Next steps</span></span>

<span data-ttu-id="e8869-148">Félicitations, vous avez créé un complément Project !</span><span class="sxs-lookup"><span data-stu-id="e8869-148">Congratulations, you've successfully created a Project add-in!</span></span> <span data-ttu-id="e8869-149">Ensuite, découvrez les fonctionnalités d’un complément Project et explorez des scénarios plus courants.</span><span class="sxs-lookup"><span data-stu-id="e8869-149">Next, learn more about the capabilities of a Project add-in and explore common scenarios.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="e8869-150">Compléments Project</span><span class="sxs-lookup"><span data-stu-id="e8869-150">Project add-ins</span></span>](../project/project-add-ins.md)
