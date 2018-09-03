<span data-ttu-id="ecece-101">Lorsqu’un tableau est tellement long que l’utilisateur doit le faire défiler pour afficher les lignes suivantes, la ligne d’en-tête peut être masquée.</span><span class="sxs-lookup"><span data-stu-id="ecece-101">When a table is long enough that a user must scroll to see some rows, the header row can scroll out of sight.</span></span> <span data-ttu-id="ecece-102">Dans cette étape du didacticiel, vous allez figer la ligne d’en-tête du tableau que vous avez créé précédemment, afin qu’elle reste visible même lorsque l’utilisateur fait défiler la feuille de calcul vers le bas.</span><span class="sxs-lookup"><span data-stu-id="ecece-102">In this step of the tutorial, you'll freeze the header row of the table that you created previously, so that it remains visible even as the user scrolls down the worksheet.</span></span> 

> [!NOTE]
> <span data-ttu-id="ecece-103">Cette page décrit une étape individuelle du didacticiel sur le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="ecece-103">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="ecece-104">Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur le complément Excel](../tutorials/excel-tutorial.yml) pour démarrer le didacticiel à partir du début.</span><span class="sxs-lookup"><span data-stu-id="ecece-104">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="freeze-the-tables-header-row"></a><span data-ttu-id="ecece-105">Figer la ligne d’en-tête du tableau</span><span class="sxs-lookup"><span data-stu-id="ecece-105">Freeze the table's header row</span></span>

1. <span data-ttu-id="ecece-106">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="ecece-106">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="ecece-107">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="ecece-107">Open the file index.html.</span></span>
3. <span data-ttu-id="ecece-108">En dessous de la balise `div` qui contient le bouton `create-chart`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="ecece-108">Below the `div` that contains the `create-chart` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="freeze-header">Freeze Header</button>            
    </div>
    ```

4. <span data-ttu-id="ecece-109">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="ecece-109">Open the app.js file.</span></span>

5. <span data-ttu-id="ecece-110">En dessous de la ligne qui attribue un gestionnaire de clic au bouton `create-chart`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="ecece-110">Below the line that assigns a click handler to the `create-chart` button, add the following code:</span></span>

    ```js
    $('#freeze-header').click(freezeHeader);
    ```

6. <span data-ttu-id="ecece-111">En dessous de la fonction `createChart`, ajoutez la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="ecece-111">Below the `createChart` function add the following function:</span></span>

    ```js
    function freezeHeader() {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to keep the header visible when the user scrolls.

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

7. <span data-ttu-id="ecece-p103">Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="ecece-p103">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="ecece-114">La collection `Worksheet.freezePanes` est un ensemble de volets de la feuille de calcul qui sont épinglés, c’est-à-dire figés, lorsque vous faites défiler la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="ecece-114">The `Worksheet.freezePanes` collection is a set of panes in the worksheet that are pinned, or frozen, in place when the worksheet is scrolled.</span></span>
   - <span data-ttu-id="ecece-p104">La méthode `freezeRows` prend comme paramètre le nombre de lignes, à partir du haut, qui doivent être figées. Nous transmettons `1` pour figer la première ligne.</span><span class="sxs-lookup"><span data-stu-id="ecece-p104">The `freezeRows` method takes as a parameter the number of rows, from the top that are to be pinned in place. We pass `1` to pin the first row in place.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="ecece-117">Tester le complément</span><span class="sxs-lookup"><span data-stu-id="ecece-117">Test the add-in</span></span>

1. <span data-ttu-id="ecece-118">Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur Ctrl+C à deux reprises pour arrêter le serveur web en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="ecece-118">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="ecece-119">Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.</span><span class="sxs-lookup"><span data-stu-id="ecece-119">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="ecece-120">Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet.</span><span class="sxs-lookup"><span data-stu-id="ecece-120">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="ecece-121">Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build.</span><span class="sxs-lookup"><span data-stu-id="ecece-121">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="ecece-122">Une fois la commande build exécutée, redémarrez le serveur.</span><span class="sxs-lookup"><span data-stu-id="ecece-122">After the build, you restart the server.</span></span> <span data-ttu-id="ecece-123">Les prochaines étapes vous permettent d’effectuer ce processus.</span><span class="sxs-lookup"><span data-stu-id="ecece-123">The next few steps carry out this process.</span></span>

1. <span data-ttu-id="ecece-124">Exécutez la commande `npm run build` pour transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par Internet Explorer (qui est utilisé en arrière-plan par Excel pour exécuter les compléments Excel).</span><span class="sxs-lookup"><span data-stu-id="ecece-124">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
2. <span data-ttu-id="ecece-125">Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.</span><span class="sxs-lookup"><span data-stu-id="ecece-125">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="ecece-126">Rechargez le volet Office en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet Office** pour rouvrir le complément.</span><span class="sxs-lookup"><span data-stu-id="ecece-126">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
6. <span data-ttu-id="ecece-127">Si le tableau est dans la feuille de calcul, supprimez-le.</span><span class="sxs-lookup"><span data-stu-id="ecece-127">If the table is in the worksheet, delete it.</span></span>
7. <span data-ttu-id="ecece-128">Dans le volet Office, sélectionnez **Créer un tableau**.</span><span class="sxs-lookup"><span data-stu-id="ecece-128">In the taskpane, choose **Create Table**.</span></span> 
8. <span data-ttu-id="ecece-129">Sélectionnez le bouton **Freeze Header**.</span><span class="sxs-lookup"><span data-stu-id="ecece-129">Choose the **Freeze Header** button.</span></span>
9. <span data-ttu-id="ecece-130">Faites suffisamment défiler la feuille de calcul vers le bas pour voir que l’en-tête du tableau est toujours visible dans la partie supérieure même lorsque les lignes du haut sont masquées.</span><span class="sxs-lookup"><span data-stu-id="ecece-130">Scroll down the worksheet enough to to see that the table header remains visible at the top even when the higher rows scroll out of sight.</span></span>

    ![Didacticiel Excel - Figer l’en-tête](../images/excel-tutorial-freeze-header.png)
