<span data-ttu-id="3cf03-101">Dans cette étape du didacticiel, vous découvrirez comment créer des contrôles de contenu de texte enrichi dans le document, puis comment insérer et remplacer du contenu dans les contrôles.</span><span class="sxs-lookup"><span data-stu-id="3cf03-101">In this step of the tutorial, you'll learn how to create Rich Text content controls in the document, and then how to insert and replace content in the controls.</span></span> 

> [!NOTE]
> <span data-ttu-id="3cf03-p101">Cette page décrit une étape individuelle d’un didacticiel sur les compléments Word. Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur les compléments Word](../tutorials/word-tutorial.yml) pour démarrer le didacticiel à partir du début.</span><span class="sxs-lookup"><span data-stu-id="3cf03-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

<span data-ttu-id="3cf03-104">Avant de commencer cette étape du didacticiel, nous vous recommandons de créer et de manipuler des contrôles de contenu de texte enrichi via l’interface utilisateur Word afin de vous familiariser avec les contrôles et leurs propriétés.</span><span class="sxs-lookup"><span data-stu-id="3cf03-104">Before you start this step of the tutorial, we recommend that you create and manipulate Rich Text content controls through the Word UI, so you can be familiar with the controls and their properties.</span></span> <span data-ttu-id="3cf03-105">Pour plus d’informations, reportez-vous à l’article [Créer des formulaires à remplir ou imprimer dans Word](https://support.office.com/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).</span><span class="sxs-lookup"><span data-stu-id="3cf03-105">For details, see [Create forms that users complete or print in Word](https://support.office.com/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).</span></span>

> [!NOTE]
> <span data-ttu-id="3cf03-106">Il existe plusieurs types de contrôles de contenu pouvant être ajoutés à un document Word via l’interface utilisateur. Toutefois, actuellement, seuls les contrôles de contenu de texte enrichi sont pris en charge par Word.js.</span><span class="sxs-lookup"><span data-stu-id="3cf03-106">There are several types of content controls that can be added to a Word document through the UI; but currently only Rich Text content controls are supported by Word.js.</span></span>


## <a name="create-a-content-control"></a><span data-ttu-id="3cf03-107">Créer un contrôle de contenu</span><span class="sxs-lookup"><span data-stu-id="3cf03-107">Create a content control</span></span>

1. <span data-ttu-id="3cf03-108">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="3cf03-108">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="3cf03-109">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="3cf03-109">Open the file index.html.</span></span>
3. <span data-ttu-id="3cf03-110">En dessous de la balise `div` qui contient le bouton `replace-text`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="3cf03-110">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="create-content-control">Create Content Control</button>            
    </div>
    ```

4. <span data-ttu-id="3cf03-111">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="3cf03-111">Open the app.js file.</span></span>

5. <span data-ttu-id="3cf03-112">Sous la ligne qui attribue un gestionnaire de clics au bouton `insert-table`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="3cf03-112">Below the line that assigns a click handler to the `insert-table` button, add the following code:</span></span>

    ```js
    $('#create-content-control').click(createContentControl);
    ```

6. <span data-ttu-id="3cf03-113">Sous la fonction `insertTable`, ajoutez la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="3cf03-113">Below the `insertTable` function, add the following function:</span></span>

    ```js
    function createContentControl() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to create a content control.

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

7. <span data-ttu-id="3cf03-p103">Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="3cf03-p103">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="3cf03-116">Ce code est destiné à intégrer l’expression « Office 365 » dans un contrôle de contenu.</span><span class="sxs-lookup"><span data-stu-id="3cf03-116">This code is intended to wrap the phrase "Office 365" in a content control.</span></span> <span data-ttu-id="3cf03-117">Cela permet d’émettre une hypothèse simplifiée selon laquelle la chaîne est présente et l’utilisateur l’a sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="3cf03-117">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>
   - <span data-ttu-id="3cf03-118">La propriété `ContentControl.title` indique le titre visible du contrôle de contenu.</span><span class="sxs-lookup"><span data-stu-id="3cf03-118">The `ContentControl.title` property specifies the visible title of the content control.</span></span> 
   - <span data-ttu-id="3cf03-119">La propriété `ContentControl.tag` indique une balise qui peut être utilisée pour obtenir une référence à un contrôle de contenu à l’aide de la méthode `ContentControlCollection.getByTag`, que vous utiliserez dans une fonction ultérieure.</span><span class="sxs-lookup"><span data-stu-id="3cf03-119">The `ContentControl.tag` property specifies an tag that can be used to get a reference to a content control using the `ContentControlCollection.getByTag` method, which you'll use in a later function.</span></span> 
   - <span data-ttu-id="3cf03-120">La propriété `ContentControl.appearance` indique l’apparence visuelle du contrôle.</span><span class="sxs-lookup"><span data-stu-id="3cf03-120">The `ContentControl.appearance` property specifies the visual look of the control.</span></span> <span data-ttu-id="3cf03-121">Utiliser la valeur « Tags » (Balises) signifie que le contrôle est intégré entre des balises de début et de fin, et que la balise de début portera le titre du contrôle de contenu.</span><span class="sxs-lookup"><span data-stu-id="3cf03-121">Using the value "Tags" means that the control will be wrapped in opening and closing tags, and the opening tag will have the content control's title.</span></span> <span data-ttu-id="3cf03-122">Les autres valeurs possibles sont « BoundingBox » (Cadre englobant) et « None » (Aucun).</span><span class="sxs-lookup"><span data-stu-id="3cf03-122">Other possible values are "BoundingBox" and "None".</span></span>
   - <span data-ttu-id="3cf03-123">La propriété `ContentControl.color` spécifie la couleur des balises ou la bordure du cadre englobant.</span><span class="sxs-lookup"><span data-stu-id="3cf03-123">The `ContentControl.color` property specifies the color of the tags or the border of the bounding box.</span></span>

    ```js
    const serviceNameRange = context.document.getSelection();
    const serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ``` 

## <a name="replace-the-content-of-the-content-control"></a><span data-ttu-id="3cf03-124">Remplacer le contenu du contrôle de contenu</span><span class="sxs-lookup"><span data-stu-id="3cf03-124">Replace the content of the content control</span></span>

1. <span data-ttu-id="3cf03-125">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="3cf03-125">Open the file index.html.</span></span>
2. <span data-ttu-id="3cf03-126">En dessous de la balise `div` qui contient le bouton `create-content-control`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="3cf03-126">Below the `div` that contains the `create-content-control` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="replace-content-in-control">Rename Service</button>            
    </div>
    ```

3. <span data-ttu-id="3cf03-127">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="3cf03-127">Open the app.js file.</span></span>

4. <span data-ttu-id="3cf03-128">Sous la ligne qui attribue un gestionnaire de clics au bouton `create-content-control`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="3cf03-128">Below the line that assigns a click handler to the `create-content-control` button, add the following code:</span></span>

    ```js
    $('#replace-content-in-control').click(replaceContentInControl);
    ```

5. <span data-ttu-id="3cf03-129">Sous la fonction `createContentControl`, ajoutez la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="3cf03-129">Below the `createContentControl` function, add the following function:</span></span>

    ```js
    function replaceContentInControl() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to replace the text in the Service Name
            //        content control.

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

7. <span data-ttu-id="3cf03-130">Remplacez `TODO1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="3cf03-130">Replace `TODO1` with the following code.</span></span> 
    > [!NOTE]
    > <span data-ttu-id="3cf03-131">La méthode `ContentControlCollection.getByTag` renvoie un `ContentControlCollection` de tous les contrôles de contenu de la balise spécifiée.</span><span class="sxs-lookup"><span data-stu-id="3cf03-131">The `ContentControlCollection.getByTag` method returns a `ContentControlCollection` of all content controls of the specified tag.</span></span> <span data-ttu-id="3cf03-132">Nous utilisons `getFirst` pour obtenir une référence au contrôle souhaité.</span><span class="sxs-lookup"><span data-stu-id="3cf03-132">We use `getFirst` to get a reference to the desired control.</span></span>

    ```js
    const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="3cf03-133">Test du complément</span><span class="sxs-lookup"><span data-stu-id="3cf03-133">Test the add-in</span></span>

1. <span data-ttu-id="3cf03-134">Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur Ctrl+C à deux reprises pour arrêter le serveur web en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="3cf03-134">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="3cf03-135">Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.</span><span class="sxs-lookup"><span data-stu-id="3cf03-135">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
     > [!NOTE]
     > <span data-ttu-id="3cf03-136">Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet.</span><span class="sxs-lookup"><span data-stu-id="3cf03-136">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="3cf03-137">Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build.</span><span class="sxs-lookup"><span data-stu-id="3cf03-137">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="3cf03-138">Après la commande build, redémarrez le serveur.</span><span class="sxs-lookup"><span data-stu-id="3cf03-138">After the build, restart the server.</span></span> <span data-ttu-id="3cf03-139">Les prochaines étapes vous permettent d’effectuer ce processus.</span><span class="sxs-lookup"><span data-stu-id="3cf03-139">The next few steps carry out this process.</span></span>
2. <span data-ttu-id="3cf03-140">Exécutez la commande `npm run build` afin de transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par tous les hôtes sur lesquels les compléments Office peuvent être exécutés.</span><span class="sxs-lookup"><span data-stu-id="3cf03-140">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="3cf03-141">Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.</span><span class="sxs-lookup"><span data-stu-id="3cf03-141">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="3cf03-142">Rechargez le volet des tâches en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet des tâches** pour rouvrir le complément.</span><span class="sxs-lookup"><span data-stu-id="3cf03-142">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="3cf03-143">Dans le volet des tâches, sélectionnez **Insérer un paragraphe** pour vous assurer qu’il existe un paragraphe contenant « Office 365 » en haut du document.</span><span class="sxs-lookup"><span data-stu-id="3cf03-143">In the taskpane, choose **Insert Paragraph** to ensure that there is a paragraph with "Office 365" at the top of the document.</span></span>
6. <span data-ttu-id="3cf03-144">Sélectionnez l’expression « Office 365 » dans le paragraphe que vous venez d’ajouter, puis sélectionnez le bouton **Créer un contrôle de contenu**.</span><span class="sxs-lookup"><span data-stu-id="3cf03-144">Select the phrase "Office 365" in the paragraph you just added, and then choose the **Create Content Control** button.</span></span> <span data-ttu-id="3cf03-145">L’expression est intégrée dans des balises nommées « Service name » (Nom de service).</span><span class="sxs-lookup"><span data-stu-id="3cf03-145">Note that the phrase is wrapped in tags labelled "Service Name".</span></span>
7. <span data-ttu-id="3cf03-146">Sélectionnez le bouton **Renommer le service** et notez que le texte du contrôle de contenu devient « Fabrikam Online Productivity Suite ».</span><span class="sxs-lookup"><span data-stu-id="3cf03-146">Choose the **Rename Service** button and note that the text of the content control changes to "Fabrikam Online Productivity Suite".</span></span>

    ![Didacticiel Word - Créer un contrôle de contenu et modifier son texte](../images/word-tutorial-content-control.png)
