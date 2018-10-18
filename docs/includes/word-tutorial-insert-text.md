<span data-ttu-id="c8bb6-101">Dans cette étape du didacticiel, vous devez tester par programme que votre complément prend en charge la version actuelle de Word de l’utilisateur, puis insérer un paragraphe dans le document.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-101">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Word, and then insert a paragraph in the document.</span></span>

> [!NOTE]
> <span data-ttu-id="c8bb6-p101">Cette page décrit une étape individuelle d’un didacticiel sur les compléments Word. Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur les compléments Word](../tutorials/word-tutorial.yml) pour démarrer le didacticiel à partir du début.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="code-the-add-in"></a><span data-ttu-id="c8bb6-104">Codage du complément</span><span class="sxs-lookup"><span data-stu-id="c8bb6-104">Code the add-in</span></span>

1. <span data-ttu-id="c8bb6-105">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="c8bb6-106">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-106">Open the file index.html.</span></span>
3. <span data-ttu-id="c8bb6-107">Remplacez `TODO1` par le codage suivant :</span><span class="sxs-lookup"><span data-stu-id="c8bb6-107">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="insert-paragraph">Insert Paragraph</button>
    ```

4. <span data-ttu-id="c8bb6-108">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-108">Open the app.js file.</span></span>
5. <span data-ttu-id="c8bb6-109">Remplacez `TODO1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-109">Replace the `TODO1` with the following code.</span></span> <span data-ttu-id="c8bb6-110">Ce code détermine si la version de Word de l’utilisateur prend en charge une version de Word.js qui inclut toutes les API utilisées dans les étapes de ce didacticiel.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-110">This code determines whether the user's version of Word supports a version of Word.js that includes all the APIs that are used in all the stages of this tutorial.</span></span> <span data-ttu-id="c8bb6-111">Dans un complément de production, utilisez le corps du bloc conditionnel pour masquer ou désactiver l’interface utilisateur appelant des API non prises en charge.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-111">In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs.</span></span> <span data-ttu-id="c8bb6-112">Cela permet à l’utilisateur de toujours utiliser les parties du complément prises en charge par sa version d’Excel.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-112">This will enable the user to still use the parts of the add-in that are supported by their version of Word.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
        console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    } 
    ```

6. <span data-ttu-id="c8bb6-113">Remplacez `TODO2` par le code suivant :</span><span class="sxs-lookup"><span data-stu-id="c8bb6-113">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#insert-paragraph').click(insertParagraph);
    ```

7. <span data-ttu-id="c8bb6-114">Remplacez `TODO3` par le code suivant :</span><span class="sxs-lookup"><span data-stu-id="c8bb6-114">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="c8bb6-115">Remarques :</span><span class="sxs-lookup"><span data-stu-id="c8bb6-115">Note the following:</span></span>
   - <span data-ttu-id="c8bb6-116">Votre logique métier Word.js est ajoutée à la fonction qui est transmise à `Word.run`.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-116">Your Word.js business logic will be added to the function that is passed to `Word.run`.</span></span> <span data-ttu-id="c8bb6-117">Cette logique n’est pas exécutée immédiatement.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-117">This logic does not execute immediately.</span></span> <span data-ttu-id="c8bb6-118">Au lieu de cela, elle est ajoutée à une file d’attente de commandes.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-118">Instead, it is added to a queue of pending commands.</span></span>
   - <span data-ttu-id="c8bb6-119">La méthode `context.sync` envoie toutes les commandes en file d’attente vers Word pour exécution.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-119">The `context.sync` method sends all queued commands to Word for execution.</span></span>
   - <span data-ttu-id="c8bb6-120">L’élément `Word.run` est suivi par un bloc `catch`.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-120">The `Word.run` is followed by a `catch` block.</span></span> <span data-ttu-id="c8bb6-121">Il s’agit d’une meilleure pratique que vous devez toujours suivre.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-121">This is a best practice that you should always follow.</span></span> 

    ```js
    function insertParagraph() {
        Word.run(function (context) {
            
            // TODO4: Queue commands to insert a paragraph into the document.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

8. <span data-ttu-id="c8bb6-p106">Remplacez `TODO4` par le code suivant. Veuillez noter les informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="c8bb6-p106">Replace `TODO4` with the following code. Note:</span></span>
   - <span data-ttu-id="c8bb6-124">Le premier paramètre de la méthode `insertParagraph` correspond au texte pour le nouveau paragraphe.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-124">The first parameter to the `insertParagraph` method is the text for the new paragraph.</span></span>
   - <span data-ttu-id="c8bb6-125">Le deuxième paramètre correspond à l’emplacement dans le corps où sera inséré le paragraphe.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-125">The second parameter is the location within the body where the paragraph will be inserted.</span></span> <span data-ttu-id="c8bb6-126">Les autres options d’insertion de paragraphe, lorsque l’objet parent est le corps, sont « Fin » et « Remplacer ».</span><span class="sxs-lookup"><span data-stu-id="c8bb6-126">Other options for insert paragraph, when the parent object is the body, are "End" and "Replace".</span></span> 

    ```js
    const docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Office 365 Click-to-Run, and Office Online.",
                            "Start");   
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="c8bb6-127">Test du complément</span><span class="sxs-lookup"><span data-stu-id="c8bb6-127">Test the add-in</span></span>

1. <span data-ttu-id="c8bb6-128">Ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-128">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
2. <span data-ttu-id="c8bb6-129">Exécutez la commande `npm run build` afin de transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par tous les hôtes sur lesquels les compléments Office peuvent être exécutés.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-129">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="c8bb6-130">Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-130">Run the command `npm start` to start a web server running on localhost.</span></span>   
4. <span data-ttu-id="c8bb6-131">Chargez une version test du complément en utilisant l’une des méthodes suivantes :</span><span class="sxs-lookup"><span data-stu-id="c8bb6-131">Sideload the add-in by using one of the following methods:</span></span>
    - <span data-ttu-id="c8bb6-132">Windows : [Chargement de version test des compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="c8bb6-132">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="c8bb6-133">Word Online : [Chargement d’une version test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="c8bb6-133">Word Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="c8bb6-134">iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="c8bb6-134">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>
5. <span data-ttu-id="c8bb6-135">Dans le menu **Accueil** de Word, sélectionnez **Afficher le volet des tâches**.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-135">On the **Home** menu of Word, select **Show Taskpane**.</span></span>
6. <span data-ttu-id="c8bb6-136">Dans le volet des tâches, sélectionnez **Insérer un paragraphe**.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-136">In the taskpane, choose **Insert Paragraph**.</span></span>
7. <span data-ttu-id="c8bb6-137">Apportez une modification au paragraphe.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-137">Make a change in the paragraph.</span></span> 
8. <span data-ttu-id="c8bb6-138">Sélectionnez à nouveau **Insérer un paragraphe**.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-138">Choose **Insert Paragraph** again.</span></span> <span data-ttu-id="c8bb6-139">Notez que le nouveau paragraphe se trouve au-dessus du paragraphe précédent, car la méthode `insertParagraph` effectue l’insertion au « début » du corps du document.</span><span class="sxs-lookup"><span data-stu-id="c8bb6-139">Note that the new paragraph is above the previous one because the `insertParagraph` method is inserting at the "start" of the document's body.</span></span>

    ![Didacticiel Word - Insérer un paragraphe](../images/word-tutorial-insert-paragraph.png)
