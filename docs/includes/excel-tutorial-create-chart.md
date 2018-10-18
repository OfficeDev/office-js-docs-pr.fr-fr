<span data-ttu-id="2240e-101">Dans cette étape du didacticiel, vous créerez un graphique à l’aide de données provenant du tableau précédemment créé, puis vous mettrez en forme le graphique.</span><span class="sxs-lookup"><span data-stu-id="2240e-101">In this step of the tutorial, you'll create a chart using data from the table that you created previously, and then format the chart.</span></span>

> [!NOTE]
> <span data-ttu-id="2240e-102">Cette page décrit une étape individuelle du didacticiel sur le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="2240e-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="2240e-103">Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur le complément Excel](../tutorials/excel-tutorial.yml) pour démarrer le didacticiel à partir du début.</span><span class="sxs-lookup"><span data-stu-id="2240e-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="chart-table-data"></a><span data-ttu-id="2240e-104">Données du tableau pour le graphique</span><span class="sxs-lookup"><span data-stu-id="2240e-104">Chart table data</span></span>

1. <span data-ttu-id="2240e-105">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="2240e-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="2240e-106">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="2240e-106">Open the file index.html.</span></span>
3. <span data-ttu-id="2240e-107">En dessous de la balise `div` qui contient le bouton `sort-table`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="2240e-107">Below the `div` that contains the `sort-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="create-chart">Create Chart</button>            
    </div>
    ```

4. <span data-ttu-id="2240e-108">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="2240e-108">Open the app.js file.</span></span>

5. <span data-ttu-id="2240e-109">Sous la ligne qui attribue un gestionnaire de clics au bouton `sort-chart`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="2240e-109">Below the line that assigns a click handler to the `sort-chart` button, add the following code:</span></span>

    ```js
    $('#create-chart').click(createChart);
    ```

6. <span data-ttu-id="2240e-110">Sous la fonction `sortTable`, ajoutez la fonction suivante.</span><span class="sxs-lookup"><span data-stu-id="2240e-110">Below the `sortTable` function add the following function.</span></span>

    ```js
    function createChart() {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to get the range of data to be charted.

            // TODO2: Queue command to create the chart and define its type.

            // TODO3: Queue commands to position and format the chart.

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

7. <span data-ttu-id="2240e-p102">Remplacez `TODO1` par le code suivant. Pour exclure la ligne d’en-tête, le code utilise la méthode `Table.getDataBodyRange` pour obtenir la plage de données à représenter sous forme de graphique à la place de la méthode `getRange`.</span><span class="sxs-lookup"><span data-stu-id="2240e-p102">Replace `TODO1` with the following code. Note that in order to exclude the header row, the code uses the `Table.getDataBodyRange` method to get the range of data you want to chart instead of the `getRange` method.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const dataRange = expensesTable.getDataBodyRange();
    ``` 

8. <span data-ttu-id="2240e-113">Remplacez `TODO2` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="2240e-113">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="2240e-114">Notez les paramètres suivants :</span><span class="sxs-lookup"><span data-stu-id="2240e-114">Note the following parameters:</span></span>
   - <span data-ttu-id="2240e-p104">Le premier paramètre transmis à la méthode `add` spécifie le type de graphique. Il en existe plusieurs dizaines de types.</span><span class="sxs-lookup"><span data-stu-id="2240e-p104">The first parameter to the `add` method specifies the type of chart. There are several dozen types.</span></span> 
   - <span data-ttu-id="2240e-117">Le deuxième paramètre spécifie la plage de données à inclure dans le graphique.</span><span class="sxs-lookup"><span data-stu-id="2240e-117">The second parameter specifies the range of data to include in the chart.</span></span> 
   - <span data-ttu-id="2240e-118">Le troisième paramètre détermine si une série de points de données provenant du tableau doit être représentée sous forme de graphique par ligne ou par colonne.</span><span class="sxs-lookup"><span data-stu-id="2240e-118">The third parameter determines whether a series of data points from the table should be charted rowwise or columnwise.</span></span> <span data-ttu-id="2240e-119">L’option `auto` demande à Excel de déterminer la meilleure méthode.</span><span class="sxs-lookup"><span data-stu-id="2240e-119">The option `auto` tells Excel to decide the best method.</span></span>

    ```js
    let chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ``` 

9. <span data-ttu-id="2240e-120">Remplacez `TODO3` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="2240e-120">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="2240e-121">La majeure partie du code est explicite.</span><span class="sxs-lookup"><span data-stu-id="2240e-121">Most of this code is self-explanatory.</span></span> <span data-ttu-id="2240e-122">Remarque :</span><span class="sxs-lookup"><span data-stu-id="2240e-122">Note:</span></span>
   - <span data-ttu-id="2240e-123">Les paramètres de la méthode `setPosition` spécifient les cellules situées en haut à gauche et en bas à droite de la zone de feuille de calcul devant contenir le graphique.</span><span class="sxs-lookup"><span data-stu-id="2240e-123">The parameters to the `setPosition` method specify the upper left and lower right cells of the worksheet area that should contain the chart.</span></span> <span data-ttu-id="2240e-124">Excel peut ajuster des éléments, tels que la largeur de ligne pour que le graphique s’affiche correctement dans l’espace attribué.</span><span class="sxs-lookup"><span data-stu-id="2240e-124">Excel can adjust things like line width to make the chart look good in the space it has been given.</span></span>
   - <span data-ttu-id="2240e-125">Une « série » est un ensemble de points de données dans une colonne du tableau.</span><span class="sxs-lookup"><span data-stu-id="2240e-125">A "series" is a set of data points from a column of the table.</span></span> <span data-ttu-id="2240e-126">Étant donné qu’il n’existe qu’une seule colonne autre que de type chaîne dans le tableau, Excel déduit que la colonne est la seule colonne de points de données pour le graphique.</span><span class="sxs-lookup"><span data-stu-id="2240e-126">Since there is only one non-string column in the table, Excel infers that the column is the only column of data points to chart.</span></span> <span data-ttu-id="2240e-127">Il interprète les autres colonnes comme des étiquettes de graphique.</span><span class="sxs-lookup"><span data-stu-id="2240e-127">It interprets the other columns as chart labels.</span></span> <span data-ttu-id="2240e-128">Par conséquent, il y aura simplement une série dans le graphique et un index 0.</span><span class="sxs-lookup"><span data-stu-id="2240e-128">So there will be just one series in the chart and it will have index 0.</span></span> <span data-ttu-id="2240e-129">Il s’agit de celle à étiqueter avec « Valeur en € ».</span><span class="sxs-lookup"><span data-stu-id="2240e-129">This is the one to label with "Value in €".</span></span> 

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in €';
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="2240e-130">Test du complément</span><span class="sxs-lookup"><span data-stu-id="2240e-130">Test the add-in</span></span>


1. <span data-ttu-id="2240e-131">Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur Ctrl+C à deux reprises pour arrêter le serveur web en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="2240e-131">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="2240e-132">Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.</span><span class="sxs-lookup"><span data-stu-id="2240e-132">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="2240e-133">Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet.</span><span class="sxs-lookup"><span data-stu-id="2240e-133">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="2240e-134">Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build.</span><span class="sxs-lookup"><span data-stu-id="2240e-134">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="2240e-135">Une fois la commande build exécutée, redémarrez le serveur.</span><span class="sxs-lookup"><span data-stu-id="2240e-135">After the build, you restart the server.</span></span> <span data-ttu-id="2240e-136">Les prochaines étapes vous permettent d’effectuer ce processus.</span><span class="sxs-lookup"><span data-stu-id="2240e-136">The next few steps carry out this process.</span></span>

1. <span data-ttu-id="2240e-137">Exécutez la commande `npm run build` pour transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par Internet Explorer (qui est utilisé en arrière-plan par Excel pour exécuter les compléments Excel).</span><span class="sxs-lookup"><span data-stu-id="2240e-137">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
2. <span data-ttu-id="2240e-138">Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.</span><span class="sxs-lookup"><span data-stu-id="2240e-138">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="2240e-139">Recharger le volet Office en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet des pages** pour rouvrir le complément.</span><span class="sxs-lookup"><span data-stu-id="2240e-139">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="2240e-140">Si pour une raison quelconque le tableau n'est pas dans la feuille de calcul ouverte, dans le volet Office, sélectionnez **Créer un tableau**, puis les boutons **Filtrer le tableau** et **Trier le tableau** dans l’ordre.</span><span class="sxs-lookup"><span data-stu-id="2240e-140">If for any reason the table is not in the open worksheet, in the taskpane, choose **Create Table** and then **Filter Table** and **Sort Table** buttons, in either order.</span></span>
6. <span data-ttu-id="2240e-141">Sélectionnez le bouton **Créer un graphique**.</span><span class="sxs-lookup"><span data-stu-id="2240e-141">Choose the **Create Chart** button.</span></span> <span data-ttu-id="2240e-142">Un graphique est créé dans lequel seules les données provenant des lignes filtrées sont incluses.</span><span class="sxs-lookup"><span data-stu-id="2240e-142">A chart is created and only the data from the rows that have been filtered are included.</span></span> <span data-ttu-id="2240e-143">Les étiquettes sur les points de données en bas sont organisées selon l’ordre de tri du graphique, à savoir les noms de marchand par ordre alphabétique inversé.</span><span class="sxs-lookup"><span data-stu-id="2240e-143">The labels on the data points across the bottom are in the sort order of the chart; that is, merchant names in reverse alphabetical order.</span></span>

    ![Didacticiel Excel - Créer un graphique](../images/excel-tutorial-create-chart.png)
