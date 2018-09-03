<span data-ttu-id="0157f-101">Dans cette étape du didacticiel, vous allez personnaliser l’interface utilisateur du volet Office.</span><span class="sxs-lookup"><span data-stu-id="0157f-101">In this step of the tutorial, you'll customize the task pane user interface (UI).</span></span>

> [!NOTE]
> <span data-ttu-id="0157f-102">Cette page décrit une étape individuelle du didacticiel sur le complément PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="0157f-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="0157f-103">Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur le complément PowerPoint](../tutorials/powerpoint-tutorial.yml) pour démarrer le didacticiel à partir du début.</span><span class="sxs-lookup"><span data-stu-id="0157f-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="customize-the-task-pane-ui"></a><span data-ttu-id="0157f-104">Personnalisation de l’interface utilisateur du volet Office</span><span class="sxs-lookup"><span data-stu-id="0157f-104">Customize the task pane UI</span></span> 

1. <span data-ttu-id="0157f-105">Dans le fichier **Home.html**, remplacez `TODO2` par le balisage suivant pour ajouter une section d’en-tête et un titre au volet Office.</span><span class="sxs-lookup"><span data-stu-id="0157f-105">In the **Home.html** file, replace `TODO2` with the following markup to add a header section and title to the task pane.</span></span> <span data-ttu-id="0157f-106">Remarque :</span><span class="sxs-lookup"><span data-stu-id="0157f-106">Note:</span></span>

    - <span data-ttu-id="0157f-107">Les styles qui commencent par `ms-` sont définis par la [structure Fabric de l’interface utilisateur Office](../design/office-ui-fabric.md), une infrastructure frontale JavaScript pour créer des expériences utilisateur pour Office et Office 365.</span><span class="sxs-lookup"><span data-stu-id="0157f-107">The styles that begin with `ms-` are defined by [Office UI Fabric](../design/office-ui-fabric.md), a JavaScript front-end framework for building user experiences for Office and Office 365.</span></span> <span data-ttu-id="0157f-108">Le fichier **Home.html** inclut une référence à la feuille de style Fabric.</span><span class="sxs-lookup"><span data-stu-id="0157f-108">The **Home.html** file includes a reference to the Fabric stylesheet.</span></span>

    ```html
    <div id="content-header">
        <div class="ms-Grid ms-bgColor-neutralPrimary">
            <div class="ms-Grid-row">
                <div class="padding ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"> <div class="ms-font-xl ms-fontColor-white ms-fontWeight-semibold">My PowerPoint Add-in</div></div>
            </div>
        </div>
    </div>
    ```

2. <span data-ttu-id="0157f-109">Dans le fichier **Home.html**, recherchez la balise **div** avec `class="footer"` et supprimez toute la balise **div** pour retirer la section de pied de page du volet Office.</span><span class="sxs-lookup"><span data-stu-id="0157f-109">In the **Home.html** file, find the **div** with `class="footer"` and delete that entire **div** to remove the footer section from the task pane.</span></span>

## <a name="test-the-add-in"></a><span data-ttu-id="0157f-110">Tester le complément</span><span class="sxs-lookup"><span data-stu-id="0157f-110">Test the add-in</span></span>

1. <span data-ttu-id="0157f-p104">À l’aide de Visual Studio, testez le complément PowerPoint en appuyant sur `F5` ou en choisissant le bouton **Démarrer** pour lancer PowerPoint avec le bouton du complément **Show Taskpane** (Afficher le volet Office) qui apparaît dans le ruban. Le complément sera hébergé localement sur IIS.</span><span class="sxs-lookup"><span data-stu-id="0157f-p104">Using Visual Studio, test the PowerPoint add-in by pressing `F5` or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![Capture d’écran de Visual Studio avec le bouton Démarrer mis en évidence](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="0157f-114">Dans PowerPoint, sélectionnez le bouton **Show Taskpane** (Afficher le volet Office) dans le ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="0157f-114">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Capture d’écran de Visual Studio avec le bouton Show Taskpane (Afficher le volet Office) mis en évidence dans le ruban Accueil](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="0157f-116">Notez que le volet Office contient désormais une section d’en-tête et un titre, et ne contient plus de section de pied de page.</span><span class="sxs-lookup"><span data-stu-id="0157f-116">Notice that the task pane now contains a header section and title, and no longer contains a footer section.</span></span>

    ![Capture d’écran du complément PowerPoint avec le bouton Insérer une image mis en évidence](../images/powerpoint-tutorial-new-task-pane-ui.png)

4. <span data-ttu-id="0157f-118">Dans Visual Studio, arrêtez le complément en appuyant sur `Shift + F5` ou en choisissant le bouton **Arrêter**.</span><span class="sxs-lookup"><span data-stu-id="0157f-118">In Visual Studio, stop the add-in by pressing `Shift + F5` or choosing the **Stop** button.</span></span> <span data-ttu-id="0157f-119">PowerPoint se ferme automatiquement lorsque le complément est arrêté.</span><span class="sxs-lookup"><span data-stu-id="0157f-119">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![Capture d’écran de Visual Studio avec le bouton Arrêter mis en évidence](../images/powerpoint-tutorial-stop.png)

