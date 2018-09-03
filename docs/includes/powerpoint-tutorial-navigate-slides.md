<span data-ttu-id="d71f1-101">Dans cette étape du didacticiel, vous allez parcourir les diapositives d’un document.</span><span class="sxs-lookup"><span data-stu-id="d71f1-101">In this step of the tutorial, you'll navigate between the slides of a document.</span></span>

> [!NOTE]
> <span data-ttu-id="d71f1-102">Cette page décrit une étape individuelle du didacticiel sur le complément PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="d71f1-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="d71f1-103">Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur le complément PowerPoint](../tutorials/powerpoint-tutorial.yml) pour démarrer le didacticiel à partir du début.</span><span class="sxs-lookup"><span data-stu-id="d71f1-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="navigate-between-slides-of-the-document"></a><span data-ttu-id="d71f1-104">Naviguer entre les diapositives du document</span><span class="sxs-lookup"><span data-stu-id="d71f1-104">Navigate between slides of the document</span></span>

1. <span data-ttu-id="d71f1-105">Dans le fichier **Home.html**, remplacez `TODO5` par le balisage suivant.</span><span class="sxs-lookup"><span data-stu-id="d71f1-105">In the **Home.html** file, replace `TODO5` with the following markup.</span></span> <span data-ttu-id="d71f1-106">Ce balisage définit les quatre boutons de navigation qui s’afficheront dans le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="d71f1-106">This markup defines the four navigation buttons that will appear within the add-in's task pane.</span></span>

    ```html
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-first-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to First Slide</span>
        <span class="ms-Button-description">Go to the first slide.</span>
    </button>
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-next-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to Next Slide</span>
        <span class="ms-Button-description">Go to the next slide.</span>
    </button>
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-previous-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to Previous Slide</span>
        <span class="ms-Button-description">Go to the previous slide.</span>
    </button>
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-last-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to Last Slide</span>
        <span class="ms-Button-description">Go to the last slide.</span>
    </button>
    ```

2. <span data-ttu-id="d71f1-107">Dans le fichier **Home.js**, remplacez `TODO8` par le code suivant pour affecter les gestionnaires d’événements pour les quatre boutons de navigation.</span><span class="sxs-lookup"><span data-stu-id="d71f1-107">In the **Home.js** file, replace `TODO8` with the following code to assign the event handlers for the four navigation buttons.</span></span>

    ```js
    $('#go-to-first-slide').click(goToFirstSlide);
    $('#go-to-next-slide').click(goToNextSlide);
    $('#go-to-previous-slide').click(goToPreviousSlide);
    $('#go-to-last-slide').click(goToLastSlide);
    ```

3. <span data-ttu-id="d71f1-108">Dans le fichier **Home.js**, remplacez `TODO9` par le code suivant pour définir les fonctions de navigation.</span><span class="sxs-lookup"><span data-stu-id="d71f1-108">In the **Home.js** file, replace `TODO9` with the following code to define the navigation functions.</span></span> <span data-ttu-id="d71f1-109">Chacune de ces fonctions utilise la fonction `goToByIdAsync` pour sélectionner une diapositive en fonction de sa position dans le document (première, dernière, précédente, suivante).</span><span class="sxs-lookup"><span data-stu-id="d71f1-109">Each of these functions uses the `goToByIdAsync` function to select a slide based upon its position in the document (first, last, previous, next).</span></span>

    ```js
    function goToFirstSlide() {
        Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToLastSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToPreviousSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToNextSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

## <a name="test-the-add-in"></a><span data-ttu-id="d71f1-110">Tester le complément</span><span class="sxs-lookup"><span data-stu-id="d71f1-110">Test the add-in</span></span>

1. <span data-ttu-id="d71f1-p104">À l’aide de Visual Studio, testez le complément en appuyant sur `F5` ou en choisissant le bouton **Démarrer** pour lancer PowerPoint avec le bouton du complément **Show Taskpane** (Afficher le volet Office) qui apparaît dans le ruban. Le complément sera hébergé localement sur IIS.</span><span class="sxs-lookup"><span data-stu-id="d71f1-p104">Using Visual Studio, test the add-in by pressing `F5` or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![Capture d’écran de Visual Studio avec le bouton Démarrer mis en évidence](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="d71f1-114">Dans PowerPoint, sélectionnez le bouton **Show Taskpane** (Afficher le volet Office) dans le ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="d71f1-114">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Capture d’écran de Visual Studio avec le bouton Show Taskpane (Afficher le volet Office) mis en évidence dans le ruban Accueil](../images/powerpoint-tutorial-show-taskpane-button.png)


3. <span data-ttu-id="d71f1-116">Utilisez le bouton **Nouvelle diapositive** dans le ruban de l’onglet **Accueil** pour ajouter deux nouvelles diapositives au document.</span><span class="sxs-lookup"><span data-stu-id="d71f1-116">Use the **New Slide** button in the ribbon of the **Home** tab to add two new slides to the document.</span></span> 

4. <span data-ttu-id="d71f1-117">Dans le volet Office, sélectionnez le bouton **Go to First Slide** (Aller à la première diapositive).</span><span class="sxs-lookup"><span data-stu-id="d71f1-117">In the task pane, choose the **Go to First Slide** button.</span></span> <span data-ttu-id="d71f1-118">La première diapositive du document est sélectionnée et affichée.</span><span class="sxs-lookup"><span data-stu-id="d71f1-118">The first slide in the document is selected and displayed.</span></span>

    ![Capture d’écran du complément PowerPoint avec le bouton Go to First Slide (Aller à la première diapositive) mis en évidence](../images/powerpoint-tutorial-go-to-first-slide.png)

5. <span data-ttu-id="d71f1-120">Dans le volet Office, sélectionnez le bouton **Go to Next Slide** (Aller à la diapositive suivante).</span><span class="sxs-lookup"><span data-stu-id="d71f1-120">In the task pane, choose the **Go to Next Slide** button.</span></span> <span data-ttu-id="d71f1-121">La diapositive suivante du document est sélectionnée et affichée.</span><span class="sxs-lookup"><span data-stu-id="d71f1-121">The next slide in the document is selected and displayed.</span></span>

    ![Capture d’écran du complément PowerPoint avec le bouton Go to Next Slide (Aller à la diapositive suivante) mis en évidence](../images/powerpoint-tutorial-go-to-next-slide.png)

6. <span data-ttu-id="d71f1-123">Dans le volet Office, sélectionnez le bouton **Go to Previous Slide** (Aller à la diapositive précédente).</span><span class="sxs-lookup"><span data-stu-id="d71f1-123">In the task pane, choose the **Go to Previous Slide** button.</span></span> <span data-ttu-id="d71f1-124">La diapositive précédente du document est sélectionnée et affichée.</span><span class="sxs-lookup"><span data-stu-id="d71f1-124">The previous slide in the document is selected and displayed.</span></span>

    ![Capture d’écran du complément PowerPoint avec le bouton Go to Previous Slide (Aller à la diapositive précédente) mis en évidence](../images/powerpoint-tutorial-go-to-previous-slide.png)

7. <span data-ttu-id="d71f1-126">Dans le volet Office, sélectionnez le bouton **Go to Last Slide** (Aller à la dernière diapositive).</span><span class="sxs-lookup"><span data-stu-id="d71f1-126">In the task pane, choose the **Go to Last Slide** button.</span></span> <span data-ttu-id="d71f1-127">La dernière diapositive du document est sélectionnée et affichée.</span><span class="sxs-lookup"><span data-stu-id="d71f1-127">The last slide in the document is selected and displayed.</span></span>

    ![Capture d’écran du complément PowerPoint avec le bouton Go to Last Slide (Aller à la dernière diapositive) mis en évidence](../images/powerpoint-tutorial-go-to-last-slide.png)

8. <span data-ttu-id="d71f1-129">Dans Visual Studio, arrêtez le complément en appuyant sur `Shift + F5` ou en choisissant le bouton **Arrêter**.</span><span class="sxs-lookup"><span data-stu-id="d71f1-129">In Visual Studio, stop the add-in by pressing `Shift + F5` or choosing the **Stop** button.</span></span> <span data-ttu-id="d71f1-130">PowerPoint se ferme automatiquement lorsque le complément est arrêté.</span><span class="sxs-lookup"><span data-stu-id="d71f1-130">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![Capture d’écran de Visual Studio avec le bouton Arrêter mis en évidence](../images/powerpoint-tutorial-stop.png)
