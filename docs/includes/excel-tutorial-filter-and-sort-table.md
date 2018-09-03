<span data-ttu-id="4283e-101">Dans cette étape du didacticiel, vous allez filtrer et trier le tableau que vous avez créé précédemment.</span><span class="sxs-lookup"><span data-stu-id="4283e-101">In this step of the tutorial, you'll filter and sort the table that you created previously.</span></span>

> [!NOTE]
> <span data-ttu-id="4283e-102">Cette page décrit une étape individuelle du didacticiel sur le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="4283e-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="4283e-103">Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur le complément Excel](../tutorials/excel-tutorial.yml) pour démarrer le didacticiel à partir du début.</span><span class="sxs-lookup"><span data-stu-id="4283e-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="filter-the-table"></a><span data-ttu-id="4283e-104">Filtrage du tableau</span><span class="sxs-lookup"><span data-stu-id="4283e-104">Filter the table</span></span>

1. <span data-ttu-id="4283e-105">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="4283e-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="4283e-106">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="4283e-106">Open the file index.html.</span></span>
3. <span data-ttu-id="4283e-107">Juste en dessous de la balise `div` qui contient le bouton `create-table`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="4283e-107">Just below the `div` that contains the `create-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="filter-table">Filter Table</button>            
    </div>
    ```

4. <span data-ttu-id="4283e-108">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="4283e-108">Open the app.js file.</span></span>

5. <span data-ttu-id="4283e-109">Juste en dessous de la ligne qui attribue un gestionnaire de clic au bouton `create-table`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="4283e-109">Just below the line that assigns a click handler to the `create-table` button, add the following code:</span></span>

    ```js
    $('#filter-table').click(filterTable);
    ```

6. <span data-ttu-id="4283e-110">Ajoutez la fonction suivante juste après la fonction `createTable`.</span><span class="sxs-lookup"><span data-stu-id="4283e-110">Just below the `createTable` function, add the following function:</span></span>

    ```js
    function filterTable() {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to filter out all expense categories except 
            //        Groceries and Education.

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

7. <span data-ttu-id="4283e-p102">Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="4283e-p102">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="4283e-113">Le code obtient tout d’abord une référence à la colonne à filtrer en transférant le nom de la colonne à la méthode `getItem`, au lieu de transmettre son index à la méthode `getItemAt` comme le fait la méthode `createTable`.</span><span class="sxs-lookup"><span data-stu-id="4283e-113">The code first gets a reference to the column that needs filtering by passing the column name to the `getItem` method, instead of passing its index to the `getItemAt` method as the `createTable` method does.</span></span> <span data-ttu-id="4283e-114">Puisque les utilisateurs peuvent déplacer des colonnes de tableau, la colonne d’un index donné peut être modifiée après la création du tableau.</span><span class="sxs-lookup"><span data-stu-id="4283e-114">Since users can move table columns, the column at a given index might change after the table is created.</span></span> <span data-ttu-id="4283e-115">Par conséquent, il est préférable d’utiliser le nom de la colonne pour obtenir une référence de la colonne.</span><span class="sxs-lookup"><span data-stu-id="4283e-115">Hence, it is safer to use the column name to get a reference to the column.</span></span> <span data-ttu-id="4283e-116">Dans le didacticiel précédent, nous avons utilisé la méthode `getItemAt` en toute sécurité, car nous l’avons utilisée dans la même méthode que celle qui crée le tableau, il n’y a donc aucune chance qu’un utilisateur ait déplacé la colonne.</span><span class="sxs-lookup"><span data-stu-id="4283e-116">We used `getItemAt` safely in the preceding tutorial, because we used it in the very same method that creates the table, so there is no chance that a user has moved the column.</span></span>
   - <span data-ttu-id="4283e-117">La méthode `applyValuesFilter` est l’une des nombreuses méthodes de filtrage sur l’objet `Filter`.</span><span class="sxs-lookup"><span data-stu-id="4283e-117">The `applyValuesFilter` method is one of several filtering methods on the `Filter` object.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);
    ``` 

## <a name="sort-the-table"></a><span data-ttu-id="4283e-118">Tri du tableau</span><span class="sxs-lookup"><span data-stu-id="4283e-118">Sort the table</span></span>

1. <span data-ttu-id="4283e-119">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="4283e-119">Open the file index.html.</span></span>
2. <span data-ttu-id="4283e-120">En dessous de la balise `div` qui contient le bouton `filter-table`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="4283e-120">Below the `div` that contains the `filter-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="sort-table">Sort Table</button>            
    </div>
    ```

3. <span data-ttu-id="4283e-121">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="4283e-121">Open the app.js file.</span></span>

4. <span data-ttu-id="4283e-122">Sous la ligne qui attribue un gestionnaire de clics au bouton `filter-table`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="4283e-122">Below the line that assigns a click handler to the `filter-table` button, add the following code:</span></span>

    ```js
    $('#sort-table').click(sortTable);
    ```

5. <span data-ttu-id="4283e-123">Ajoutez la fonction suivante après la fonction `filterTable`.</span><span class="sxs-lookup"><span data-stu-id="4283e-123">Below the `filterTable` function add the following function.</span></span>

    ```js
    function sortTable() {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to sort the table by Merchant name.

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

7. <span data-ttu-id="4283e-p104">Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="4283e-p104">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="4283e-126">Le code crée un tableau d’objets `SortField` qui ne comporte qu’un seul membre puisque le complément ne trie que la colonne Merchant.</span><span class="sxs-lookup"><span data-stu-id="4283e-126">The code creates an array of `SortField` objects which has just one member since the add-in only sorts on the Merchant column.</span></span>
   - <span data-ttu-id="4283e-127">La propriété `key` d’un objet `SortField` est l’index de la colonne à trier qui part de zéro.</span><span class="sxs-lookup"><span data-stu-id="4283e-127">The `key` property of a `SortField` object is the zero-based index of the column to sort-on.</span></span>
   - <span data-ttu-id="4283e-128">Le membre `sort` d’un objet `Table` est un objet `TableSort`, et non une méthode.</span><span class="sxs-lookup"><span data-stu-id="4283e-128">The `sort` member of a `Table` is a `TableSort` object, not a method.</span></span> <span data-ttu-id="4283e-129">Les objets `SortField` sont transmis à la méthode `apply` de l’objet `TableSort`.</span><span class="sxs-lookup"><span data-stu-id="4283e-129">The `SortField`s are passed the `TableSort` object's `apply` method.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const sortFields = [
        { 
            key: 1,            // Merchant column
            ascending: false,
        }
    ];

    expensesTable.sort.apply(sortFields);
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="4283e-130">Test du complément</span><span class="sxs-lookup"><span data-stu-id="4283e-130">Test the add-in</span></span>

1. <span data-ttu-id="4283e-131">Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur Ctrl+C à deux reprises pour arrêter le serveur web en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="4283e-131">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="4283e-132">Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.</span><span class="sxs-lookup"><span data-stu-id="4283e-132">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="4283e-133">Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet.</span><span class="sxs-lookup"><span data-stu-id="4283e-133">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="4283e-134">Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build.</span><span class="sxs-lookup"><span data-stu-id="4283e-134">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="4283e-135">Une fois la commande build exécutée, redémarrez le serveur.</span><span class="sxs-lookup"><span data-stu-id="4283e-135">After the build, you restart the server.</span></span> <span data-ttu-id="4283e-136">Les prochaines étapes vous permettent d’effectuer ce processus.</span><span class="sxs-lookup"><span data-stu-id="4283e-136">The next few steps carry out this process.</span></span>

1. <span data-ttu-id="4283e-137">Exécutez la commande `npm run build` pour transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par Internet Explorer (qui est utilisé en arrière-plan par Excel pour exécuter les compléments Excel).</span><span class="sxs-lookup"><span data-stu-id="4283e-137">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
2. <span data-ttu-id="4283e-138">Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.</span><span class="sxs-lookup"><span data-stu-id="4283e-138">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="4283e-139">Rechargez le volet Office en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet Office** pour rouvrir le complément.</span><span class="sxs-lookup"><span data-stu-id="4283e-139">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="4283e-140">Si, pour une raison quelconque, le tableau ne se trouve pas dans la feuille de calcul ouverte, dans le volet Office, sélectionnez **Créer un tableau**.</span><span class="sxs-lookup"><span data-stu-id="4283e-140">If for any reason the table is not in the open worksheet, in the taskpane, choose **Create Table**.</span></span> 
6. <span data-ttu-id="4283e-141">Choisissez les boutons **Filtrer le tableau** et **Trier le tableau** dans n’importe quel ordre.</span><span class="sxs-lookup"><span data-stu-id="4283e-141">Choose the **Filter Table** and **Sort Table** buttons, in either order.</span></span>

    ![Didacticiel Excel - Filtrer et trier un tableau](../images/excel-tutorial-filter-and-sort-table.png)
