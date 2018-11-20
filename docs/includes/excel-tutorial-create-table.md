<span data-ttu-id="f78d0-101">Dans cette étape du didacticiel, vous vérifiez à l’aide de programme que votre complément prend en charge la version actuelle Excel de l’utilisateur, vous ajoutez un tableau à une feuille de calcul, vous renseignez le tableau avec des données et vous le mettez en forme.</span><span class="sxs-lookup"><span data-stu-id="f78d0-101">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Excel, add a table to a worksheet, populate the table with data, and format it.</span></span>

> [!NOTE]
> <span data-ttu-id="f78d0-102">Cette page décrit une étape individuelle du didacticiel sur le complément Excel.</span><span class="sxs-lookup"><span data-stu-id="f78d0-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="f78d0-103">Si vous êtes arrivé à cette page via les résultats du moteur de recherche ou d’un autre lien direct, accédez à la page d’introduction du [didacticiel sur le complément Excel](../tutorials/excel-tutorial.yml) pour démarrer le didacticiel à partir du début.</span><span class="sxs-lookup"><span data-stu-id="f78d0-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="code-the-add-in"></a><span data-ttu-id="f78d0-104">Codage du complément</span><span class="sxs-lookup"><span data-stu-id="f78d0-104">Code the add-in</span></span>

1. <span data-ttu-id="f78d0-105">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="f78d0-105">Open the project in your code editor.</span></span>
2. <span data-ttu-id="f78d0-106">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="f78d0-106">Open the file index.html.</span></span>
3. <span data-ttu-id="f78d0-107">Remplacez `TODO1` par le codage suivant :</span><span class="sxs-lookup"><span data-stu-id="f78d0-107">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="create-table">Create Table</button>
    ```

4. <span data-ttu-id="f78d0-108">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="f78d0-108">Open the app.js file.</span></span>
5. <span data-ttu-id="f78d0-109">Remplacez `TODO1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="f78d0-109">Replace the `TODO1` with the following code.</span></span> <span data-ttu-id="f78d0-110">Ce code détermine si la version Excel de l’utilisateur prend en charge une version d’Excel.js qui inclut toutes les API utilisées dans cette série de didacticiels.</span><span class="sxs-lookup"><span data-stu-id="f78d0-110">This code determines whether the user's version of Excel supports a version of Excel.js that includes all the APIs that this series of tutorials will use.</span></span> <span data-ttu-id="f78d0-111">Dans un complément de production, utilisez le corps du bloc conditionnel pour masquer ou désactiver l’interface utilisateur appelant des API non prises en charge.</span><span class="sxs-lookup"><span data-stu-id="f78d0-111">In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs.</span></span> <span data-ttu-id="f78d0-112">Cela permet à l’utilisateur de toujours utiliser les parties du complément prises en charge par leur version d’Excel.</span><span class="sxs-lookup"><span data-stu-id="f78d0-112">This will enable the user to still make use of the parts of the add-in that are supported by their version of Excel.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }
    ```

6. <span data-ttu-id="f78d0-113">Remplacez `TODO2` par le code suivant :</span><span class="sxs-lookup"><span data-stu-id="f78d0-113">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#create-table').click(createTable);
    ```

7. <span data-ttu-id="f78d0-114">Remplacez `TODO3` par le code suivant :</span><span class="sxs-lookup"><span data-stu-id="f78d0-114">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="f78d0-115">Remarques :</span><span class="sxs-lookup"><span data-stu-id="f78d0-115">Note the following:</span></span>
   - <span data-ttu-id="f78d0-116">Votre logique métier Excel.js est ajoutée à la fonction qui est transmise à `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="f78d0-116">Your Excel.js business logic will be added to the function that is passed to `Excel.run`.</span></span> <span data-ttu-id="f78d0-117">Cette logique n’est pas exécutée immédiatement.</span><span class="sxs-lookup"><span data-stu-id="f78d0-117">This logic does not execute immediately.</span></span> <span data-ttu-id="f78d0-118">Au lieu de cela, elle est ajoutée à une file d’attente de commandes.</span><span class="sxs-lookup"><span data-stu-id="f78d0-118">Instead, it is added to a queue of pending commands.</span></span>
   - <span data-ttu-id="f78d0-119">La méthode `context.sync` envoie toutes les commandes en file d’attente vers Excel pour exécution.</span><span class="sxs-lookup"><span data-stu-id="f78d0-119">The `context.sync` method sends all queued commands to Excel for execution.</span></span>
   - <span data-ttu-id="f78d0-120">L’élément `Excel.run` est suivi par un bloc `catch`.</span><span class="sxs-lookup"><span data-stu-id="f78d0-120">The `Excel.run` is followed by a `catch` block.</span></span> <span data-ttu-id="f78d0-121">Il s’agit d’une meilleure pratique que vous devez toujours suivre.</span><span class="sxs-lookup"><span data-stu-id="f78d0-121">This is a best practice that you should always follow.</span></span> 

    ```js
    function createTable() {
        Excel.run(function (context) {

            // TODO4: Queue table creation logic here.

            // TODO5: Queue commands to populate the table with data.

            // TODO6: Queue commands to format the table.

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

8. <span data-ttu-id="f78d0-p106">Remplacez `TODO4` par le code suivant. Remarque :</span><span class="sxs-lookup"><span data-stu-id="f78d0-p106">Replace `TODO4` with the following code. Note:</span></span>
   - <span data-ttu-id="f78d0-124">Le code crée un tableau à l’aide de la méthode `add` de collection de tableau d’une feuille de calcul, qui existe toujours même si elle est vide.</span><span class="sxs-lookup"><span data-stu-id="f78d0-124">The code creates a table by using `add` method of a worksheet's table collection, which always exists even if it is empty.</span></span> <span data-ttu-id="f78d0-125">Il s’agit de la méthode standard de création d’objets Excel.js.</span><span class="sxs-lookup"><span data-stu-id="f78d0-125">This is the standard way that Excel.js objects are created.</span></span> <span data-ttu-id="f78d0-126">Il n’existe aucune API pour le constructeur de classe API. De plus, vous n’utilisez jamais d’opérateur `new` pour créer un objet Excel.</span><span class="sxs-lookup"><span data-stu-id="f78d0-126">There are no class constructor APIs, and you never use a `new` operator to create an Excel object.</span></span> <span data-ttu-id="f78d0-127">Au lieu de cela, vous l’ajoutez à un objet de la collection parent.</span><span class="sxs-lookup"><span data-stu-id="f78d0-127">Instead, you add to a parent collection object.</span></span>
   - <span data-ttu-id="f78d0-128">Le premier paramètre de la méthode `add` est la plage comprenant uniquement la ligne supérieure du tableau, et non la plage entière utilisée en fin de compte par le tableau.</span><span class="sxs-lookup"><span data-stu-id="f78d0-128">The first parameter of the `add` method is the range of only the top row of the table, not the entire range the table will ultimately use.</span></span> <span data-ttu-id="f78d0-129">La raison est que lorsque le complément remplit les lignes de données (dans l’étape suivante), il ajoute de nouvelles lignes au tableau au lieu d’écrire des valeurs dans les cellules des lignes existantes.</span><span class="sxs-lookup"><span data-stu-id="f78d0-129">This is because when the add-in populates the data rows (in the next step), it will add new rows to the table instead of writing values to the cells of existing rows.</span></span> <span data-ttu-id="f78d0-130">Il s’agit d’un modèle plus courant, car le nombre de lignes contenues dans un tableau est souvent inconnu lorsque le tableau est créé.</span><span class="sxs-lookup"><span data-stu-id="f78d0-130">This is a more common pattern because the number of rows that a table will have is often not known when the table is created.</span></span>
   - <span data-ttu-id="f78d0-131">Les noms de tableau doivent être uniques dans l’ensemble du classeur, pas uniquement dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="f78d0-131">Table names must be unique across the entire workbook, not just the worksheet.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ```

9. <span data-ttu-id="f78d0-p109">Remplacez `TODO5` par le code suivant. Remarque :</span><span class="sxs-lookup"><span data-stu-id="f78d0-p109">Replace `TODO5` with the following code. Note:</span></span>
   - <span data-ttu-id="f78d0-134">Les valeurs de cellule d’une plage sont définies avec un tableau de tableaux.</span><span class="sxs-lookup"><span data-stu-id="f78d0-134">The cell values of a range are set with an array of arrays.</span></span>
   - <span data-ttu-id="f78d0-135">Les nouvelles lignes sont créées dans un tableau en appelant la méthode `add` de collection de ligne du tableau.</span><span class="sxs-lookup"><span data-stu-id="f78d0-135">New rows are created in a table by calling the `add` method of the table's row collection.</span></span> <span data-ttu-id="f78d0-136">Vous pouvez ajouter plusieurs lignes dans un seul appel de `add` en incluant plusieurs tableaux de valeurs de cellule dans le tableau parent transmis en tant que deuxième paramètre.</span><span class="sxs-lookup"><span data-stu-id="f78d0-136">You can add multiple rows in a single call of `add` by including multiple cell value arrays in the parent array that is passed as the second parameter.</span></span>

    ```js
    expensesTable.getHeaderRowRange().values =
        [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add at the end*/, [
        ["1/1/2017", "The Phone Company", "Communications", "120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
        ["1/11/2017", "Bellows College", "Education", "350.1"],
        ["1/15/2017", "Trey Research", "Other", "135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
    ]);
    ```

10. <span data-ttu-id="f78d0-p111">Remplacez `TODO6` par le code suivant. Remarque :</span><span class="sxs-lookup"><span data-stu-id="f78d0-p111">Replace `TODO6` with the following code. Note:</span></span>
   - <span data-ttu-id="f78d0-139">Le code recherche une référence à la colonne **Amount** en transmettant son index de base zéro à la méthode `getItemAt` de collection de colonnes du tableau.</span><span class="sxs-lookup"><span data-stu-id="f78d0-139">The code gets a reference to the **Amount** column by passing its zero-based index to the `getItemAt` method of the table's column collection.</span></span>

     > [!NOTE]
     > <span data-ttu-id="f78d0-140">Les objets de collection Excel.js, tels que `TableCollection`, `WorksheetCollection` et `TableColumnCollection` ont une propriété `items` qui correspond à un tableau de types d’objet enfant, comme `Table` ou `Worksheet` ou `TableColumn` ; mais un objet `*Collection` n’est pas lui-même un tableau.</span><span class="sxs-lookup"><span data-stu-id="f78d0-140">Excel.js collection objects, such as `TableCollection`, `WorksheetCollection`, and `TableColumnCollection` have an `items` property that is an array of the child object types, such as `Table` or `Worksheet` or `TableColumn`; but a `*Collection` object is not itself an array.</span></span>

   - <span data-ttu-id="f78d0-141">Le code définit ensuite la plage de la colonne **Amount** sous la forme Euros à la deuxième décimale.</span><span class="sxs-lookup"><span data-stu-id="f78d0-141">The code then formats the range of the **Amount** column as Euros to the second decimal.</span></span> 
   - <span data-ttu-id="f78d0-142">Enfin, il s’assure que la largeur des colonnes et la hauteur des lignes sont assez grandes pour contenir l’élément de données le plus long (ou le plus haut).</span><span class="sxs-lookup"><span data-stu-id="f78d0-142">Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item.</span></span> <span data-ttu-id="f78d0-143">Notez que le code doit rechercher des objets `Range` à mettre en forme.</span><span class="sxs-lookup"><span data-stu-id="f78d0-143">Notice that the code must get `Range` objects to format.</span></span> <span data-ttu-id="f78d0-144">Les objets `TableColumn` et `TableRow` n’ont pas de propriétés de mise en forme.</span><span class="sxs-lookup"><span data-stu-id="f78d0-144">`TableColumn` and `TableRow` objects do not have format properties.</span></span>

        ```js
        expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
        expensesTable.getRange().format.autofitColumns();
        expensesTable.getRange().format.autofitRows();
        ```

## <a name="test-the-add-in"></a><span data-ttu-id="f78d0-145">Test du complément</span><span class="sxs-lookup"><span data-stu-id="f78d0-145">Test the add-in</span></span>

1. <span data-ttu-id="f78d0-146">Ouvrez une fenêtre Git Bash ou une invite système activée par Node.JS, et accédez au dossier **Démarrer** du projet.</span><span class="sxs-lookup"><span data-stu-id="f78d0-146">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
2. <span data-ttu-id="f78d0-147">Exécutez la commande `npm run build` pour transpiler votre code source ES6 sur une version antérieure de JavaScript prise en charge par Internet Explorer (qui est utilisée en arrière-plan par Excel pour exécuter les compléments Excel).</span><span class="sxs-lookup"><span data-stu-id="f78d0-147">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
3. <span data-ttu-id="f78d0-148">Exécutez la commande `npm start` pour démarrer un serveur web exécuté sur un hôte local.</span><span class="sxs-lookup"><span data-stu-id="f78d0-148">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="f78d0-149">Chargez une version test du complément en utilisant l’une des méthodes suivantes :</span><span class="sxs-lookup"><span data-stu-id="f78d0-149">Sideload the add-in by using one of the following methods:</span></span>
    - <span data-ttu-id="f78d0-150">Windows : [Chargement de versions test de compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="f78d0-150">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="f78d0-151">Excel Online : [Chargement de versions test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="f78d0-151">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="f78d0-152">iPad et Mac : [Chargement de versions test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="f78d0-152">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>
5. <span data-ttu-id="f78d0-153">Dans le menu **Accueil**, sélectionnez **Afficher le volet Office**.</span><span class="sxs-lookup"><span data-stu-id="f78d0-153">On the **Home** menu, choose **Show Taskpane**.</span></span>
6. <span data-ttu-id="f78d0-154">Dans le volet Office, sélectionnez **Créer un tableau**.</span><span class="sxs-lookup"><span data-stu-id="f78d0-154">In the taskpane, choose **Create Table**.</span></span>

    ![Didacticiel Excel - Créer un tableau](../images/excel-tutorial-create-table.png)
