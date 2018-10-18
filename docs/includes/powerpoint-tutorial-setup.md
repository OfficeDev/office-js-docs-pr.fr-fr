<span data-ttu-id="23546-101">Vous commencerez ce didacticiel par la configuration de votre projet de développement.</span><span class="sxs-lookup"><span data-stu-id="23546-101">You'll begin this tutorial by setting up your development project.</span></span> 

> [!NOTE]
> <span data-ttu-id="23546-102">Cette page décrit une étape individuelle du didacticiel sur le complément PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="23546-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="23546-103">Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur le complément PowerPoint](../tutorials/powerpoint-tutorial.yml) pour démarrer le didacticiel à partir du début.</span><span class="sxs-lookup"><span data-stu-id="23546-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="23546-104">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="23546-104">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

## <a name="setup"></a><span data-ttu-id="23546-105">Installation</span><span class="sxs-lookup"><span data-stu-id="23546-105">Setup</span></span>

<span data-ttu-id="23546-106">Dans ce didacticiel, vous allez créer un complément à l’aide de Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="23546-106">In this tutorial, you'll create an add-in using Visual Studio.</span></span>

### <a name="create-the-add-in-project"></a><span data-ttu-id="23546-107">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="23546-107">Create the add-in project</span></span>

1. <span data-ttu-id="23546-108">Dans la barre de menu de Visual Studio, choisissez successivement **Fichier** > **Nouveau** > **Projet**.</span><span class="sxs-lookup"><span data-stu-id="23546-108">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="23546-109">Dans la liste des types de projet sous **Visual C#** ou **Visual Basic**, développez **Office/SharePoint**, choisissez **Compléments**, puis **Complément web PowerPoint** pour le type de projet.</span><span class="sxs-lookup"><span data-stu-id="23546-109">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **PowerPoint Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="23546-110">Nommez le projet **HelloWorld**, puis sélectionnez le bouton **OK**.</span><span class="sxs-lookup"><span data-stu-id="23546-110">Name the project **HelloWorld**, and then choose the **OK** button.</span></span>

4. <span data-ttu-id="23546-111">Dans la fenêtre de la boîte de dialogue **Créer un complément Office**, choisissez **Ajouter de nouvelles fonctionnalités à PowerPoint**, puis sélectionnez **Terminer** pour créer le projet.</span><span class="sxs-lookup"><span data-stu-id="23546-111">In the **Create Office Add-in** dialog window, choose **Add new functionalities to PowerPoint**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="23546-p102">Visual Studio crée une solution et ses deux projets apparaissent dans l’**explorateur de solutions**. Le fichier **Home.html** s’ouvre dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="23546-p102">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

     ![Didacticiel PowerPoint - Fenêtre de l’explorateur de solutions Visual Studio qui affiche les 2 projets dans la solution HelloWorld](../images/powerpoint-tutorial-solution-explorer.png)

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="23546-115">Explorer la solution Visual Studio</span><span class="sxs-lookup"><span data-stu-id="23546-115">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-code"></a><span data-ttu-id="23546-116">Mise à jour du code</span><span class="sxs-lookup"><span data-stu-id="23546-116">Update code</span></span> 

<span data-ttu-id="23546-117">Modifiez le code de complément comme suit pour créer la structure que vous utiliserez pour implémenter la fonctionnalité de complément dans les étapes suivantes de ce didacticiel.</span><span class="sxs-lookup"><span data-stu-id="23546-117">Edit the add-in code as follows, to create the framework that you'll use to implement add-in functionality in subsequent steps of this tutorial.</span></span>

1. <span data-ttu-id="23546-118">**Home.html** spécifie le code HTML qui s’affichera dans le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="23546-118">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="23546-119">Dans **Home.html**, localisez la balise **div** avec `id="content-main"`, remplacez l’intégralité de la balise **div** avec le balisage suivant et enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="23546-119">In **Home.html**, find the **div** with `id="content-main"`, replace that entire **div** with the following markup, and save the file.</span></span>

    ```html
    <!-- TODO2: Create the content-header div. -->
    <div id="content-main">
        <div class="padding">
            <!-- TODO1: Create the insert-image button. -->
            <!-- TODO3: Create the insert-text button. -->
            <!-- TODO4: Create the get-slide-metadata button. -->
            <!-- TODO5: Create the go-to-slide buttons. -->
        </div>
    </div>
    ```

2. <span data-ttu-id="23546-120">Ouvrez le fichier **Home.js** à la racine du projet d’application web.</span><span class="sxs-lookup"><span data-stu-id="23546-120">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="23546-121">Ce fichier spécifie le script pour le complément.</span><span class="sxs-lookup"><span data-stu-id="23546-121">This file specifies the script for the add-in.</span></span> <span data-ttu-id="23546-122">Remplacez tout le contenu par le code suivant, puis enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="23546-122">Replace the entire contents with the following code and save the file.</span></span>

    ```javascript
    (function () {
        "use strict";

        var messageBanner;

        Office.initialize = function (reason) {
            $(document).ready(function () {
                // Initialize the FabricUI notification mechanism and hide it
                var element = document.querySelector('.ms-MessageBanner');
                messageBanner = new fabric.MessageBanner(element);
                messageBanner.hideBanner();

                // TODO1: Assign event handler for insert-image button.
                // TODO4: Assign event handler for insert-text button.
                // TODO6: Assign event handler for get-slide-metadata button.
                // TODO8: Assign event handlers for the four navigation buttons.
            });
        };

        // TODO2: Define the insertImage function. 

        // TODO3: Define the insertImageFromBase64String function.

        // TODO5: Define the insertText function.

        // TODO7: Define the getSlideMetadata function.

        // TODO9: Define the navigation functions.

        // Helper function for displaying notifications
        function showNotification(header, content) {
            $("#notification-header").text(header);
            $("#notification-body").text(content);
            messageBanner.showBanner();
            messageBanner.toggleExpansion();
        }
    })();
    ```
