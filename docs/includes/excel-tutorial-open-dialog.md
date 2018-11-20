<span data-ttu-id="caf3b-101">Dans cette étape finale du didacticiel, vous allez ouvrir une boîte de dialogue dans votre complément, transmettre un message du processus de boîte de dialogue au processus de volet Office et fermer la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="caf3b-101">In this final step of the tutorial, you'll open a dialog in your add-in, pass a message from the dialog process to the task pane process, and close the dialog.</span></span> <span data-ttu-id="caf3b-102">Les boîtes de dialogue des compléments Office sont *non modales* : un utilisateur peut continuer à interagir à la fois avec le document dans l’application Office hôte et avec la page hôte dans le volet Office.</span><span class="sxs-lookup"><span data-stu-id="caf3b-102">Office Add-in dialogs are *nonmodal*: a user can continue to interact with both the document in the host Office application and with the host page in the task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="caf3b-103">Cette page décrit une étape individuelle du didacticiel sur le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="caf3b-103">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="caf3b-104">Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur le complément Excel](../tutorials/excel-tutorial.yml) pour démarrer le didacticiel à partir du début.</span><span class="sxs-lookup"><span data-stu-id="caf3b-104">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="create-the-dialog-page"></a><span data-ttu-id="caf3b-105">Création de la page de boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="caf3b-105">Create the dialog page</span></span>

1. <span data-ttu-id="caf3b-106">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="caf3b-106">Open the project in your code editor.</span></span>
2. <span data-ttu-id="caf3b-107">Créez un fichier à la racine du projet (où se trouve le fichier index.html) et nommez-le popup.html.</span><span class="sxs-lookup"><span data-stu-id="caf3b-107">Create a file in the root of the project (where index.html is) called popup.html.</span></span>
3. <span data-ttu-id="caf3b-p103">Ajoutez le balisage suivant au fichier popup.html. Remarque :</span><span class="sxs-lookup"><span data-stu-id="caf3b-p103">Add the following markup to popup.html. Note:</span></span>
   - <span data-ttu-id="caf3b-110">La page comporte un champ `<input>`, dans lequel l’utilisateur entrera son nom, et un bouton qui permet d’envoyer le nom à la page dans le volet Office où il sera affiché.</span><span class="sxs-lookup"><span data-stu-id="caf3b-110">The page has a `<input>` where the user will enter his or her name and a button that will send the name to the page in the task pane where it will be displayed.</span></span>
   - <span data-ttu-id="caf3b-111">Le balisage charge un script appelé popup.js que vous allez créer dans une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="caf3b-111">The markup loads a script called popup.js that you will create in a later step.</span></span>
   - <span data-ttu-id="caf3b-112">Il charge également la bibliothèque Office.JS et jQuery, car ils seront utilisés dans popup.js.</span><span class="sxs-lookup"><span data-stu-id="caf3b-112">It also loads the Office.JS library and jQuery because they will be used in popup.js.</span></span>

    ```html
    <!DOCTYPE html>
    <html>
        <head lang="en">
            <title>Dialog for My Office Add-in</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">

            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.min.css" />
            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.components.css" />
            <link rel="stylesheet" href="app.css" />

            <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
            <script type="text/javascript" src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.2.1.min.js"></script>
            <script type="text/javascript" src="popup.js"></script>

        </head>
        <body style="display:flex;flex-direction:column;align-items:center;justify-content:center">
            <div class="padding">
                <p class="ms-font-xl">ENTER YOUR NAME</p>
            </div>
            <div class="padding">
                <input id="name-box" type="text"/>
            </div>
            <div class="padding">
                <button id="ok-button" class="ms-Button">OK</button>
            </div>
        </body>
    </html>
    ```

4. <span data-ttu-id="caf3b-113">Créez un fichier à la racine du projet et nommez-le popup.js.</span><span class="sxs-lookup"><span data-stu-id="caf3b-113">Create a file in the root of the project called popup.js.</span></span>
5. <span data-ttu-id="caf3b-p104">Ajoutez le code suivant au fichier popup.js. Remarque :</span><span class="sxs-lookup"><span data-stu-id="caf3b-p104">Add the following code to popup.js. Note:</span></span>
   - <span data-ttu-id="caf3b-116">*Toutes les pages qui appellent des API dans la bibliothèque Office.JS doivent affecter une fonction à la propriété `Office.initialize`.*</span><span class="sxs-lookup"><span data-stu-id="caf3b-116">*Every page that calls APIs in the Office.JS library must assign a function to the `Office.initialize` property.*</span></span> <span data-ttu-id="caf3b-117">Si aucune initialisation n’est nécessaire, la fonction peut avoir un corps vide, mais la propriété ne doit pas être laissée indéfinie, affectée à null ni à une valeur qui n’est pas une fonction.</span><span class="sxs-lookup"><span data-stu-id="caf3b-117">If no initialization is needed, then the function can have an empty body, but the property must not be left undefined, assigned to null or to a non-function value.</span></span> <span data-ttu-id="caf3b-118">Pour voir un exemple, affichez le fichier app.js à la racine du projet.</span><span class="sxs-lookup"><span data-stu-id="caf3b-118">For an example, see the app.js file in the project root.</span></span> <span data-ttu-id="caf3b-119">Le code qui exécute l’affectation doit être exécuté avant tout appel à Office.JS ; l’affectation se trouve donc dans un fichier de script chargé par la page, comme dans ce cas.</span><span class="sxs-lookup"><span data-stu-id="caf3b-119">The code that makes the assignment must run before any calls to Office.JS; hence the assignment is in a script file that is loaded by the page, as it is in this case.</span></span>
   - <span data-ttu-id="caf3b-p106">La fonction `ready` jQuery est appelée à l’intérieur de la méthode `initialize`. Une règle quasi-universelle veut que le code de chargement, d’initialisation ou d’amorçage des autres bibliothèques JavaScript se trouve à l’intérieur de la fonction `Office.initialize`.</span><span class="sxs-lookup"><span data-stu-id="caf3b-p106">The jQuery `ready` function is called inside the `initialize` method. It is an almost universal rule that the loading, initializing, or bootstrapping code of other JavaScript libraries should be inside the `Office.initialize` function.</span></span>

    ```js
    (function () {
    "use strict";

        Office.initialize = function() {
            $(document).ready(function () {  

                // TODO1: Assign handler to the OK button.

            });
        }

        // TODO2: Create the OK button handler

    }());
    ```

6. <span data-ttu-id="caf3b-122">Remplacez `TODO1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="caf3b-122">Replace `TODO1` with the following code:</span></span> <span data-ttu-id="caf3b-123">Vous allez créer la fonction `sendStringToParentPage` à l’étape suivante.</span><span class="sxs-lookup"><span data-stu-id="caf3b-123">You'll create the `sendStringToParentPage` function in the next step.</span></span>

    ```js
    $('#ok-button').click(sendStringToParentPage);
    ```

7. <span data-ttu-id="caf3b-124">Remplacez `TODO2` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="caf3b-124">Replace `TODO2` with the following code:</span></span> <span data-ttu-id="caf3b-125">La méthode `messageParent` transmet son paramètre à la page parent, qui est, dans ce cas, la page dans le volet Office.</span><span class="sxs-lookup"><span data-stu-id="caf3b-125">The `messageParent` method passes its parameter to the parent page, in this case, the page in the task pane.</span></span> <span data-ttu-id="caf3b-126">Le paramètre peut être une valeur booléenne ou une chaîne qui inclut tous les éléments qui peuvent être sérialisés en tant que chaîne, au format XML ou JSON.</span><span class="sxs-lookup"><span data-stu-id="caf3b-126">The parameter can be a boolean or a string, which includes anything that can be serialized as a string, such as XML or JSON.</span></span>

    ```js
    function sendStringToParentPage() {
        var userName = $('#name-box').val();
        Office.context.ui.messageParent(userName);
    }
    ```

8. <span data-ttu-id="caf3b-127">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="caf3b-127">Save the file.</span></span>

   > [!NOTE]
   > <span data-ttu-id="caf3b-128">Le fichier popup.html et le fichier popup.js qu’il charge s’exécutent dans un processus Internet Explorer entièrement séparé à partir du volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="caf3b-128">The popup.html file, and the popup.js file that it loads, run in an entirely separate Internet Explorer process from the add-in's task pane.</span></span> <span data-ttu-id="caf3b-129">Si le popup.js était transpilé dans le même fichier bundle.js en tant que fichier app.js, le complément devrait charger deux copies du fichier bundle.js, ce qui irait à l’encontre de l’objectif de groupement.</span><span class="sxs-lookup"><span data-stu-id="caf3b-129">If the popup.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="caf3b-130">En outre, le fichier popup.js ne contient pas de code JavaScript car Internet Explorer ne prend pas en charge ce type de code.</span><span class="sxs-lookup"><span data-stu-id="caf3b-130">In addition, the popup.js file does not contain any JavaScript that is unsupported by IE.</span></span> <span data-ttu-id="caf3b-131">C’est pour ces deux raisons que ce complément ne transpile pas le fichier popup.js du tout.</span><span class="sxs-lookup"><span data-stu-id="caf3b-131">For these two reasons, this add-in does not transpile the popup.js file at all.</span></span>


## <a name="open-the-dialog-from-the-task-pane"></a><span data-ttu-id="caf3b-132">Ouverture de la boîte de dialogue à partir du volet Office</span><span class="sxs-lookup"><span data-stu-id="caf3b-132">Open the dialog from the task pane</span></span>

1. <span data-ttu-id="caf3b-133">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="caf3b-133">Open the file index.html.</span></span>
2. <span data-ttu-id="caf3b-134">Sous la balise `div` qui contient le bouton `freeze-header`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="caf3b-134">Below the `div` that contains the `freeze-header` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="open-dialog">Open Dialog</button>
    </div>
    ```

3. <span data-ttu-id="caf3b-135">La boîte de dialogue invitera l’utilisateur à saisir son nom et transmettra ce nom au volet Office.</span><span class="sxs-lookup"><span data-stu-id="caf3b-135">The dialog will prompt the user to enter a name and pass the user's name to the task pane.</span></span> <span data-ttu-id="caf3b-136">Le volet Office s’affichera dans une étiquette.</span><span class="sxs-lookup"><span data-stu-id="caf3b-136">The task pane will display it in a label.</span></span> <span data-ttu-id="caf3b-137">Juste en dessous de la balise `div` que vous venez d’ajouter, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="caf3b-137">Immediately below the `div` that you just added, add the following markup:</span></span>

    ```html
    <div class="padding">
        <label id="user-name"></label>
    </div>
    ```

4. <span data-ttu-id="caf3b-138">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="caf3b-138">Open the app.js file.</span></span>

5. <span data-ttu-id="caf3b-139">Sous la ligne qui attribue un gestionnaire de clics au bouton `freeze-header`, ajoutez le code suivant.</span><span class="sxs-lookup"><span data-stu-id="caf3b-139">Below the line that assigns a click handler to the `freeze-header` button, add the following code.</span></span> <span data-ttu-id="caf3b-140">Vous allez créer la méthode `openDialog` à une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="caf3b-140">You'll create the `openDialog` method in a later step.</span></span>

    ```js
    $('#open-dialog').click(openDialog);
    ```

6. <span data-ttu-id="caf3b-p112">Ajoutez la déclaration suivante sous la fonction `freezeHeader`. Cette variable est utilisée pour conserver un objet dans le contexte d’exécution de la page parent qui agit en tant qu’intermédiaire pour le contexte d’exécution de la page de boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="caf3b-p112">Below the `freezeHeader` function add the following declaration. This variable is used to hold an object in the parent page's execution context that acts as an intermediator to the dialog page's execution context.</span></span>

    ```js
    let dialog = null;
    ```

7. <span data-ttu-id="caf3b-143">Sous la déclaration de la balise `dialog`, ajoutez la fonction suivante.</span><span class="sxs-lookup"><span data-stu-id="caf3b-143">Below the declaration of `dialog`, add the following function.</span></span> <span data-ttu-id="caf3b-144">Le plus important à remarquer à propos de ce code est ce qui ne s’y trouve *pas* : il n’y a aucun appel de `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="caf3b-144">The important thing to notice about this code is what is *not* there: there is no call of `Excel.run`.</span></span> <span data-ttu-id="caf3b-145">Cela est dû au fait que l’API d’ouverture de boîte de dialogue est partagée par tous les hôtes Office, elle fait donc partie de l’API commune JavaScript Office, pas de l’API spécifique d’Excel.</span><span class="sxs-lookup"><span data-stu-id="caf3b-145">This is because the API to open a dialog is shared among all Office hosts, so it is part of the Office JavaScript Common API, not the Excel-specific API.</span></span>

    ```js
    function openDialog() {
        // TODO1: Call the Office Shared API that opens a dialog
    }
    ```

8. <span data-ttu-id="caf3b-p114">Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="caf3b-p114">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="caf3b-148">La méthode `displayDialogAsync` ouvre une boîte de dialogue au centre de l’écran.</span><span class="sxs-lookup"><span data-stu-id="caf3b-148">The `displayDialogAsync` method opens a dialog in the center of the screen.</span></span>
   - <span data-ttu-id="caf3b-149">Le premier paramètre est l’URL de la page à ouvrir.</span><span class="sxs-lookup"><span data-stu-id="caf3b-149">The first parameter is the URL of the page to open.</span></span>
   - <span data-ttu-id="caf3b-p115">Le deuxième paramètre transmet les options. `height` et `width` sont des pourcentages de la taille de la fenêtre de l’application Office.</span><span class="sxs-lookup"><span data-stu-id="caf3b-p115">The second parameter passes options. `height` and `width` are percentages of the size of the Office application's window.</span></span>

    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},

        // TODO2: Add callback parameter.
    );
    ```

## <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a><span data-ttu-id="caf3b-152">Traitement du message à partir de la boîte de dialogue et fermeture de la boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="caf3b-152">Process the message from the dialog and close the dialog</span></span>

1. <span data-ttu-id="caf3b-p116">Continuez dans le fichier app.js et remplacez `TODO2` par le code suivant. Remarque :</span><span class="sxs-lookup"><span data-stu-id="caf3b-p116">Continue in the app.js file, and replace `TODO2` with the following code. Note:</span></span>
   - <span data-ttu-id="caf3b-155">Le rappel est exécuté immédiatement après que la boîte de dialogue s’est ouverte correctement et avant que l’utilisateur ait pris une quelconque action dans la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="caf3b-155">The callback is executed immediately after the dialog successfully opens and before the user has taken any action in the dialog.</span></span>
   - <span data-ttu-id="caf3b-156">`result.value` représente l’objet qui agit comme un intermédiaire entre les contextes d’exécution des pages parent et de boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="caf3b-156">The `result.value` is the object that acts as a kind of middleman between the execution contexts of the parent and dialog pages.</span></span>
   - <span data-ttu-id="caf3b-157">La fonction `processMessage` sera créée à une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="caf3b-157">The `processMessage` function will be created in a later step.</span></span> <span data-ttu-id="caf3b-158">Ce gestionnaire traitera toutes les valeurs envoyées par la page de boîte de dialogue avec les appels de la fonction `messageParent`.</span><span class="sxs-lookup"><span data-stu-id="caf3b-158">This handler will process any values that are sent from the dialog page with calls of the `messageParent` function.</span></span>

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. <span data-ttu-id="caf3b-159">Sous la fonction `openDialog`, ajoutez la fonction suivante.</span><span class="sxs-lookup"><span data-stu-id="caf3b-159">Below the `openDialog` function, add the following function.</span></span>

    ```js
    function processMessage(arg) {
        $('#user-name').text(arg.message);
        dialog.close();
    }
    ```

## <a name="test-the-add-in"></a><span data-ttu-id="caf3b-160">Test du complément</span><span class="sxs-lookup"><span data-stu-id="caf3b-160">Test the add-in</span></span>

1. <span data-ttu-id="caf3b-161">Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur Ctrl+C à deux reprises pour arrêter le serveur web en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="caf3b-161">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="caf3b-162">Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.</span><span class="sxs-lookup"><span data-stu-id="caf3b-162">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="caf3b-163">Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet.</span><span class="sxs-lookup"><span data-stu-id="caf3b-163">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="caf3b-164">Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build.</span><span class="sxs-lookup"><span data-stu-id="caf3b-164">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="caf3b-165">Une fois la commande build exécutée, redémarrez le serveur.</span><span class="sxs-lookup"><span data-stu-id="caf3b-165">After the build, you restart the server.</span></span> <span data-ttu-id="caf3b-166">Les prochaines étapes vous permettent d’effectuer ce processus.</span><span class="sxs-lookup"><span data-stu-id="caf3b-166">The next few steps carry out this process.</span></span>

1. <span data-ttu-id="caf3b-167">Exécutez la commande `npm run build` pour transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par Internet Explorer (qui est utilisé en arrière-plan par Excel pour exécuter les compléments Excel).</span><span class="sxs-lookup"><span data-stu-id="caf3b-167">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
2. <span data-ttu-id="caf3b-168">Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.</span><span class="sxs-lookup"><span data-stu-id="caf3b-168">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="caf3b-169">Recharger le volet Office en le fermant, puis, dans le menu **Accueil**, sélectionnez **Afficher le volet des pages** pour rouvrir le complément.</span><span class="sxs-lookup"><span data-stu-id="caf3b-169">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
6. <span data-ttu-id="caf3b-170">Sélectionnez le bouton **Boîte de dialogue Ouvrir** dans le volet Office.</span><span class="sxs-lookup"><span data-stu-id="caf3b-170">Choose the **Open Dialog** button in the task pane.</span></span>
7. <span data-ttu-id="caf3b-171">Lorsque la boîte de dialogue est ouverte, faites-la glisser et redimensionnez-la.</span><span class="sxs-lookup"><span data-stu-id="caf3b-171">While the dialog is open, drag it and resize it.</span></span> <span data-ttu-id="caf3b-172">Vous pouvez interagir avec la feuille de calcul et appuyer sur les autres boutons du volet Office.</span><span class="sxs-lookup"><span data-stu-id="caf3b-172">Note that you can interact with the worksheet and press other buttons on the taskpane.</span></span> <span data-ttu-id="caf3b-173">Pour autant, vous ne pouvez pas lancer une deuxième boîte de dialogue à partir de la même page de volet Office.</span><span class="sxs-lookup"><span data-stu-id="caf3b-173">But you cannot launch a second dialog from the same task pane page.</span></span>
8. <span data-ttu-id="caf3b-174">Dans la boîte de dialogue, entrez un nom et appuyez sur **OK**.</span><span class="sxs-lookup"><span data-stu-id="caf3b-174">In the dialog, enter a name and choose **OK**.</span></span> <span data-ttu-id="caf3b-175">Ce nom apparaît sur le volet Office et la boîte de dialogue se ferme.</span><span class="sxs-lookup"><span data-stu-id="caf3b-175">The name appears on the task pane and the dialog closes.</span></span>
9. <span data-ttu-id="caf3b-176">Si vous le souhaitez, vous pouvez commenter la ligne `dialog.close();` dans la fonction `processMessage`.</span><span class="sxs-lookup"><span data-stu-id="caf3b-176">Optionally, comment out the line `dialog.close();` in the `processMessage` function.</span></span> <span data-ttu-id="caf3b-177">Ensuite, répétez les étapes de cette section.</span><span class="sxs-lookup"><span data-stu-id="caf3b-177">Then repeat the steps of this section.</span></span> <span data-ttu-id="caf3b-178">La boîte de dialogue reste ouverte et vous pouvez modifier le nom.</span><span class="sxs-lookup"><span data-stu-id="caf3b-178">The dialog stays open and you can change the name.</span></span> <span data-ttu-id="caf3b-179">Vous pouvez la fermer manuellement en appuyant sur la croix (**X**) en haut à droite.</span><span class="sxs-lookup"><span data-stu-id="caf3b-179">You can close it manually by pressing the **X** button in the upper right corner.</span></span>

    ![Didacticiel Excel - Boîte de dialogue](../images/excel-tutorial-dialog-open.png)
