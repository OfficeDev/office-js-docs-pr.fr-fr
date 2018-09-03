<span data-ttu-id="cedc5-101">Dans cette étape du didacticiel, vous allez récupérer les métadonnées de la diapositive sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="cedc5-101">In this step of the tutorial, you'll retrieve metadata for the selected slide.</span></span>

> [!NOTE]
> <span data-ttu-id="cedc5-102">Cette page décrit une étape individuelle du didacticiel sur le complément PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="cedc5-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="cedc5-103">Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur le complément PowerPoint](../tutorials/powerpoint-tutorial.yml) pour démarrer le didacticiel à partir du début.</span><span class="sxs-lookup"><span data-stu-id="cedc5-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="get-slide-metadata"></a><span data-ttu-id="cedc5-104">Obtenir les métadonnées des diapositives</span><span class="sxs-lookup"><span data-stu-id="cedc5-104">Get slide metadata</span></span>

1. <span data-ttu-id="cedc5-105">Dans le fichier **Home.html**, remplacez `TODO4` par le balisage suivant.</span><span class="sxs-lookup"><span data-stu-id="cedc5-105">In the **Home.html** file, replace `TODO4` with the following markup.</span></span> <span data-ttu-id="cedc5-106">Ce balisage définit le bouton **Get Slide Metadata** (Obtenir les métadonnées de la diapositive) qui s’affichera dans le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="cedc5-106">This markup defines the **Get Slide Metadata** button that will appear within the add-in's task pane.</span></span>

    ```html
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="get-slide-metadata">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Get Slide Metadata</span>
        <span class="ms-Button-description">Gets metadata for the selected slide(s).</span>
    </button>
    ```

2. <span data-ttu-id="cedc5-107">Dans le fichier **Home.js**, remplacez `TODO6` par le code suivant pour attribuer le gestionnaire d’événements pour le bouton **Get Slide Metadata** (Obtenir les métadonnées de la diapositive).</span><span class="sxs-lookup"><span data-stu-id="cedc5-107">In the **Home.js** file, replace `TODO6` with the following code to assign the event handler for the **Get Slide Metadata** button.</span></span>

    ```js
    $('#get-slide-metadata').click(getSlideMetadata);
    ```

3. <span data-ttu-id="cedc5-108">Dans le fichier **Home.js**, remplacez `TODO7` par le code suivant pour définir la fonction **getSlideMetadata**.</span><span class="sxs-lookup"><span data-stu-id="cedc5-108">In the **Home.js** file, replace `TODO7` with the following code to define the **getSlideMetadata** function.</span></span> <span data-ttu-id="cedc5-109">Cette fonction extrait les métadonnées pour la ou les diapositives sélectionnée(s), et les écrit dans une fenêtre de boîte de dialogue contextuelle dans le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="cedc5-109">This function retrieves metadata for the selected slide(s) and writes it to a popup dialog window within the add-in task pane.</span></span>

    ```js
    function getSlideMetadata() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                } else {
                    showNotification("Metadata for selected slide(s):", JSON.stringify(asyncResult.value), null, 2);
                }
            }
        );
    }
    ```

## <a name="test-the-add-in"></a><span data-ttu-id="cedc5-110">Tester le complément</span><span class="sxs-lookup"><span data-stu-id="cedc5-110">Test the add-in</span></span>

1. <span data-ttu-id="cedc5-p104">À l’aide de Visual Studio, testez le complément en appuyant sur `F5` ou en choisissant le bouton **Démarrer** pour lancer PowerPoint avec le bouton du complément **Show Taskpane** (Afficher le volet Office) qui apparaît dans le ruban. Le complément sera hébergé localement sur IIS.</span><span class="sxs-lookup"><span data-stu-id="cedc5-p104">Using Visual Studio, test the add-in by pressing `F5` or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![Capture d’écran de Visual Studio avec le bouton Démarrer mis en évidence](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="cedc5-114">Dans PowerPoint, sélectionnez le bouton **Show Taskpane** (Afficher le volet Office) dans le ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="cedc5-114">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Capture d’écran de Visual Studio avec le bouton Show Taskpane (Afficher le volet Office) mis en évidence dans le ruban Accueil](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="cedc5-116">Dans le volet Office, sélectionnez le bouton **Get Slide Metadata** (Obtenir les métadonnées de la diapositive) pour obtenir les métadonnées pour la diapositive sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="cedc5-116">In the task pane, choose the **Get Slide Metadata** button to get the metadata for the selected slide.</span></span> <span data-ttu-id="cedc5-117">Les métadonnées de la diapositive sont écrites dans la fenêtre de boîte de dialogue contextuelle en bas du volet Office.</span><span class="sxs-lookup"><span data-stu-id="cedc5-117">The slide metadata is written to the popup dialog window at the bottom of the task pane.</span></span> <span data-ttu-id="cedc5-118">Dans ce cas, le tableau `slides` figurant dans les métadonnées JSON contient un objet qui spécifie les éléments `id`, `title` et `index` de la diapositive sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="cedc5-118">In this case, the `slides` array within the JSON metadata contains one object that specifies the `id`, `title`, and `index` of the selected slide.</span></span> <span data-ttu-id="cedc5-119">Si plusieurs diapositives étaient sélectionnées lorsque vous avez récupéré les métadonnées des diapositives, le tableau `slides` figurant dans les métadonnées JSON contiendrait un objet pour chaque diapositive sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="cedc5-119">If multiple slides had been selected when you retrieved slide metadata, the `slides` array within the JSON metadata would contain one object for each selected slide.</span></span>

    ![Capture d’écran du complément PowerPoint avec le bouton Get Slide Metadata (Obtenir les métadonnées de la diapositive) mis en évidence](../images/powerpoint-tutorial-get-slide-metadata.png)

4. <span data-ttu-id="cedc5-121">Dans Visual Studio, arrêtez le complément en appuyant sur `Shift + F5` ou en choisissant le bouton **Arrêter**.</span><span class="sxs-lookup"><span data-stu-id="cedc5-121">In Visual Studio, stop the add-in by pressing `Shift + F5` or choosing the **Stop** button.</span></span> <span data-ttu-id="cedc5-122">PowerPoint se ferme automatiquement lorsque le complément est arrêté.</span><span class="sxs-lookup"><span data-stu-id="cedc5-122">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![Capture d’écran de Visual Studio avec le bouton Arrêter mis en évidence](../images/powerpoint-tutorial-stop.png)
