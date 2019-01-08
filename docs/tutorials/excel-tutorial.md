---
title: Didacticiel sur le complément Excel
description: Dans ce didacticiel, vous allez développer un complément Excel qui crée, remplit, filtre et trie un tableau, crée un graphique, fige un en-tête de tableau, protège une feuille de calcul et ouvre une boîte de dialogue.
ms.date: 12/31/2018
ms.topic: tutorial
ms.openlocfilehash: fe4350f5f3fdbe34250c1739c7651a1dde1e28ef
ms.sourcegitcommit: 3007bf57515b0811ff98a7e1518ecc6fc9462276
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/04/2019
ms.locfileid: "27724946"
---
# <a name="tutorial-create-an-excel-task-pane-add-in"></a><span data-ttu-id="ff084-103">Didacticiel : Créer un complément de volet de tâches de Excel</span><span class="sxs-lookup"><span data-stu-id="ff084-103">Tutorial: Create an Excel task pane add-in</span></span>

<span data-ttu-id="ff084-104">Dans ce tutoriel, vous allez créer un complément de volet de tâches Excel qui:</span><span class="sxs-lookup"><span data-stu-id="ff084-104">In this tutorial, you'll create an Excel task pane add-in that:</span></span>

> [!div class="checklist"]
> * <span data-ttu-id="ff084-105">Crée un tableau</span><span class="sxs-lookup"><span data-stu-id="ff084-105">Creates a new table.</span></span>
> * <span data-ttu-id="ff084-106">Filtres et tris un tableau</span><span class="sxs-lookup"><span data-stu-id="ff084-106">Filters and sorts a table</span></span>
> * <span data-ttu-id="ff084-107">Crée un graphique (Chart)</span><span class="sxs-lookup"><span data-stu-id="ff084-107">Creates a new chart.</span></span>
> * <span data-ttu-id="ff084-108">Figer une en-tête de tableau</span><span class="sxs-lookup"><span data-stu-id="ff084-108">Freezes a table header</span></span>
> * <span data-ttu-id="ff084-109">Protège une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="ff084-109">Protects a worksheet</span></span>
> * <span data-ttu-id="ff084-110">Ouvrir une boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="ff084-110">Opens a dialog</span></span>

## <a name="prerequisites"></a><span data-ttu-id="ff084-111">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ff084-111">Prerequisites</span></span>

<span data-ttu-id="ff084-112">Pour utiliser ce didacticiel, les logiciels suivants doivent être installés.</span><span class="sxs-lookup"><span data-stu-id="ff084-112">To use this tutorial, you need to have the following installed.</span></span> 

- <span data-ttu-id="ff084-113">Excel 2016, version 1711 (Démarrer en un clic version 8730.1000) ou version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="ff084-113">Excel 2016, version 1711 (Build 8730.1000 Click-to-Run) or later.</span></span> <span data-ttu-id="ff084-114">Vous devrez peut-être participer au programme Office Insider pour obtenir cette version.</span><span class="sxs-lookup"><span data-stu-id="ff084-114">You might need to be an Office Insider to get this version.</span></span> <span data-ttu-id="ff084-115">Pour plus d’informations, reportez-vous à [Participez au programme Office Insider](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="ff084-115">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span>

- [<span data-ttu-id="ff084-116">Node</span><span class="sxs-lookup"><span data-stu-id="ff084-116">Node</span></span>](https://nodejs.org/en/) 

- <span data-ttu-id="ff084-117">[Git Bash](https://git-scm.com/downloads) (ou un autre client Git)</span><span class="sxs-lookup"><span data-stu-id="ff084-117">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

## <a name="create-your-add-in-project"></a><span data-ttu-id="ff084-118">Créer votre projet de complément</span><span class="sxs-lookup"><span data-stu-id="ff084-118">Create your add-in project</span></span>

<span data-ttu-id="ff084-119">Procédez comme suit pour créer le projet de complément Excel que vous souhaitez utiliser comme base pour ce didacticiel.</span><span class="sxs-lookup"><span data-stu-id="ff084-119">Complete the following steps to create the Excel add-in project that you'll use as the basis for this tutorial.</span></span>

1. <span data-ttu-id="ff084-120">Clonez le référentiel GitHub du [didacticiel sur les compléments Excel](https://github.com/OfficeDev/Excel-Add-in-Tutorial).</span><span class="sxs-lookup"><span data-stu-id="ff084-120">Clone the GitHub repository [Excel Add-in Tutorial](https://github.com/OfficeDev/Excel-Add-in-Tutorial).</span></span>

2. <span data-ttu-id="ff084-121">Ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.</span><span class="sxs-lookup"><span data-stu-id="ff084-121">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

3. <span data-ttu-id="ff084-122">Exécutez la commande `npm install` pour installer les outils et les bibliothèques répertoriées dans le fichier package.json.</span><span class="sxs-lookup"><span data-stu-id="ff084-122">Run the command `npm install` to install the tools and libraries listed in the package.json file.</span></span> 

4. <span data-ttu-id="ff084-123">Effectuez les étapes décrites dans la rubrique relative à l’[ajout de certificats auto-signés comme certificat racine approuvé](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) pour approuver le certificat pour le système d’exploitation de votre ordinateur de développement.</span><span class="sxs-lookup"><span data-stu-id="ff084-123">Carry out the steps in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to trust the certificate for your development computer's operating system.</span></span>

## <a name="create-a-table"></a><span data-ttu-id="ff084-124">Créer un tableau</span><span class="sxs-lookup"><span data-stu-id="ff084-124">Create a table</span></span>

<span data-ttu-id="ff084-125">Dans cette étape du didacticiel, vous vérifiez à l’aide de programme que votre complément prend en charge la version actuelle Excel de l’utilisateur, vous ajoutez un tableau à une feuille de calcul, vous renseignez le tableau avec des données et vous le mettez en forme.</span><span class="sxs-lookup"><span data-stu-id="ff084-125">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Excel, add a table to a worksheet, populate the table with data, and format it.</span></span>

### <a name="code-the-add-in"></a><span data-ttu-id="ff084-126">Codage du complément</span><span class="sxs-lookup"><span data-stu-id="ff084-126">Code the add-in</span></span>

1. <span data-ttu-id="ff084-127">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="ff084-127">Open the project in your code editor.</span></span>

2. <span data-ttu-id="ff084-128">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="ff084-128">Open the file index.html.</span></span>

3. <span data-ttu-id="ff084-129">Remplacez `TODO1` par le codage suivant :</span><span class="sxs-lookup"><span data-stu-id="ff084-129">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="create-table">Create Table</button>
    ```

4. <span data-ttu-id="ff084-130">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="ff084-130">Open the app.js file.</span></span>

5. <span data-ttu-id="ff084-131">Remplacez `TODO1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="ff084-131">Replace the `TODO1` with the following code.</span></span> <span data-ttu-id="ff084-132">Ce code détermine si la version Excel de l’utilisateur prend en charge une version d’Excel.js qui inclut toutes les API utilisées dans cette série de didacticiels.</span><span class="sxs-lookup"><span data-stu-id="ff084-132">This code determines whether the user's version of Excel supports a version of Excel.js that includes all the APIs that this series of tutorials will use.</span></span> <span data-ttu-id="ff084-133">Dans un complément de production, utilisez le corps du bloc conditionnel pour masquer ou désactiver l’interface utilisateur appelant des API non prises en charge.</span><span class="sxs-lookup"><span data-stu-id="ff084-133">In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs.</span></span> <span data-ttu-id="ff084-134">Cela permet à l’utilisateur de toujours utiliser les parties du complément prises en charge par leur version d’Excel.</span><span class="sxs-lookup"><span data-stu-id="ff084-134">This will enable the user to still make use of the parts of the add-in that are supported by their version of Excel.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }
    ```

6. <span data-ttu-id="ff084-135">Remplacez `TODO2` par le code suivant :</span><span class="sxs-lookup"><span data-stu-id="ff084-135">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#create-table').click(createTable);
    ```

7. <span data-ttu-id="ff084-136">Remplacez `TODO3` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="ff084-136">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="ff084-137">Remarque :</span><span class="sxs-lookup"><span data-stu-id="ff084-137">Note:</span></span>

   - <span data-ttu-id="ff084-138">Votre logique métier Excel.js est ajoutée à la fonction qui est transmise à `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="ff084-138">Your Excel.js business logic will be added to the function that is passed to `Excel.run`.</span></span> <span data-ttu-id="ff084-139">Cette logique n’est pas exécutée immédiatement.</span><span class="sxs-lookup"><span data-stu-id="ff084-139">This logic does not execute immediately.</span></span> <span data-ttu-id="ff084-140">Au lieu de cela, elle est ajoutée à une file d’attente de commandes.</span><span class="sxs-lookup"><span data-stu-id="ff084-140">Instead, it is added to a queue of pending commands.</span></span>

   - <span data-ttu-id="ff084-141">La méthode `context.sync` envoie toutes les commandes en file d’attente vers Excel pour exécution.</span><span class="sxs-lookup"><span data-stu-id="ff084-141">The `context.sync` method sends all queued commands to Excel for execution.</span></span>

   - <span data-ttu-id="ff084-142">L’élément `Excel.run` est suivi par un bloc `catch`.</span><span class="sxs-lookup"><span data-stu-id="ff084-142">The `Excel.run` is followed by a `catch` block.</span></span> <span data-ttu-id="ff084-143">Il s’agit d’une meilleure pratique que vous devez toujours suivre.</span><span class="sxs-lookup"><span data-stu-id="ff084-143">This is a best practice that you should always follow.</span></span> 

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

8. <span data-ttu-id="ff084-p106">Remplacez `TODO4` par le code suivant. Remarque :</span><span class="sxs-lookup"><span data-stu-id="ff084-p106">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="ff084-146">Le code crée un tableau à l’aide de la méthode `add` de collection de tableau d’une feuille de calcul, qui existe toujours même si elle est vide.</span><span class="sxs-lookup"><span data-stu-id="ff084-146">The code creates a table by using `add` method of a worksheet's table collection, which always exists even if it is empty.</span></span> <span data-ttu-id="ff084-147">Il s’agit de la méthode standard de création d’objets Excel.js.</span><span class="sxs-lookup"><span data-stu-id="ff084-147">This is the standard way that Excel.js objects are created.</span></span> <span data-ttu-id="ff084-148">Il n’existe aucune API pour le constructeur de classe API. De plus, vous n’utilisez jamais d’opérateur `new` pour créer un objet Excel.</span><span class="sxs-lookup"><span data-stu-id="ff084-148">There are no class constructor APIs, and you never use a `new` operator to create an Excel object.</span></span> <span data-ttu-id="ff084-149">Au lieu de cela, vous l’ajoutez à un objet de la collection parent.</span><span class="sxs-lookup"><span data-stu-id="ff084-149">Instead, you add to a parent collection object.</span></span>

   - <span data-ttu-id="ff084-150">Le premier paramètre de la méthode `add` est la plage comprenant uniquement la ligne supérieure du tableau, et non la plage entière utilisée en fin de compte par le tableau.</span><span class="sxs-lookup"><span data-stu-id="ff084-150">The first parameter of the `add` method is the range of only the top row of the table, not the entire range the table will ultimately use.</span></span> <span data-ttu-id="ff084-151">La raison est que lorsque le complément remplit les lignes de données (dans l’étape suivante), il ajoute de nouvelles lignes au tableau au lieu d’écrire des valeurs dans les cellules des lignes existantes.</span><span class="sxs-lookup"><span data-stu-id="ff084-151">This is because when the add-in populates the data rows (in the next step), it will add new rows to the table instead of writing values to the cells of existing rows.</span></span> <span data-ttu-id="ff084-152">Il s’agit d’un modèle plus courant, car le nombre de lignes contenues dans un tableau est souvent inconnu lorsque le tableau est créé.</span><span class="sxs-lookup"><span data-stu-id="ff084-152">This is a more common pattern because the number of rows that a table will have is often not known when the table is created.</span></span>

   - <span data-ttu-id="ff084-153">Les noms de tableau doivent être uniques dans l’ensemble du classeur, pas uniquement dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="ff084-153">Table names must be unique across the entire workbook, not just the worksheet.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ```

9. <span data-ttu-id="ff084-p109">Remplacez `TODO5` par le code suivant. Remarque :</span><span class="sxs-lookup"><span data-stu-id="ff084-p109">Replace `TODO5` with the following code. Note:</span></span>

   - <span data-ttu-id="ff084-156">Les valeurs de cellule d’une plage sont définies avec un tableau de tableaux.</span><span class="sxs-lookup"><span data-stu-id="ff084-156">The cell values of a range are set with an array of arrays.</span></span>

   - <span data-ttu-id="ff084-157">Les nouvelles lignes sont créées dans un tableau en appelant la méthode `add` de collection de ligne du tableau.</span><span class="sxs-lookup"><span data-stu-id="ff084-157">New rows are created in a table by calling the `add` method of the table's row collection.</span></span> <span data-ttu-id="ff084-158">Vous pouvez ajouter plusieurs lignes dans un seul appel de `add` en incluant plusieurs tableaux de valeurs de cellule dans le tableau parent transmis en tant que deuxième paramètre.</span><span class="sxs-lookup"><span data-stu-id="ff084-158">You can add multiple rows in a single call of `add` by including multiple cell value arrays in the parent array that is passed as the second parameter.</span></span>

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

10. <span data-ttu-id="ff084-p111">Remplacez `TODO6` par le code suivant. Remarque :</span><span class="sxs-lookup"><span data-stu-id="ff084-p111">Replace `TODO6` with the following code. Note:</span></span>

   - <span data-ttu-id="ff084-161">Le code recherche une référence à la colonne **Amount** en transmettant son index de base zéro à la méthode `getItemAt` de collection de colonnes du tableau.</span><span class="sxs-lookup"><span data-stu-id="ff084-161">The code gets a reference to the **Amount** column by passing its zero-based index to the `getItemAt` method of the table's column collection.</span></span>

     > [!NOTE]
     > <span data-ttu-id="ff084-162">Les objets de collection Excel.js, tels que `TableCollection`, `WorksheetCollection` et `TableColumnCollection` ont une propriété `items` qui correspond à un tableau de types d’objet enfant, comme `Table` ou `Worksheet` ou `TableColumn` ; mais un objet `*Collection` n’est pas lui-même un tableau.</span><span class="sxs-lookup"><span data-stu-id="ff084-162">Excel.js collection objects, such as `TableCollection`, `WorksheetCollection`, and `TableColumnCollection` have an `items` property that is an array of the child object types, such as `Table` or `Worksheet` or `TableColumn`; but a `*Collection` object is not itself an array.</span></span>

   - <span data-ttu-id="ff084-163">Le code définit ensuite la plage de la colonne **Amount** sous la forme Euros à la deuxième décimale.</span><span class="sxs-lookup"><span data-stu-id="ff084-163">The code then formats the range of the **Amount** column as Euros to the second decimal.</span></span> 

   - <span data-ttu-id="ff084-164">Enfin, il s’assure que la largeur des colonnes et la hauteur des lignes sont assez grandes pour contenir l’élément de données le plus long (ou le plus haut).</span><span class="sxs-lookup"><span data-stu-id="ff084-164">Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item.</span></span> <span data-ttu-id="ff084-165">Notez que le code doit rechercher des objets `Range` à mettre en forme.</span><span class="sxs-lookup"><span data-stu-id="ff084-165">Notice that the code must get `Range` objects to format.</span></span> <span data-ttu-id="ff084-166">Les objets `TableColumn` et `TableRow` n’ont pas de propriétés de mise en forme.</span><span class="sxs-lookup"><span data-stu-id="ff084-166">`TableColumn` and `TableRow` objects do not have format properties.</span></span>

        ```js
        expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
        expensesTable.getRange().format.autofitColumns();
        expensesTable.getRange().format.autofitRows();
        ```

### <a name="test-the-add-in"></a><span data-ttu-id="ff084-167">Test du complément</span><span class="sxs-lookup"><span data-stu-id="ff084-167">Test the add-in</span></span>

1. <span data-ttu-id="ff084-168">Ouvrez une fenêtre Git Bash ou une invite système activée par Node.JS, et accédez au dossier **Démarrer** du projet.</span><span class="sxs-lookup"><span data-stu-id="ff084-168">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

2. <span data-ttu-id="ff084-169">Exécutez la commande `npm run build` pour transpiler votre code source ES6 sur une version antérieure de JavaScript prise en charge par Internet Explorer (qui est utilisée en arrière-plan par Excel pour exécuter les compléments Excel).</span><span class="sxs-lookup"><span data-stu-id="ff084-169">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="ff084-170">Exécutez la commande `npm start` pour démarrer un serveur web exécuté sur un hôte local.</span><span class="sxs-lookup"><span data-stu-id="ff084-170">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="ff084-171">Chargez une version test du complément en utilisant l’une des méthodes suivantes :</span><span class="sxs-lookup"><span data-stu-id="ff084-171">Sideload the add-in by using one of the following methods:</span></span>

    - <span data-ttu-id="ff084-172">Windows : [Chargement de versions test de compléments Office sur Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="ff084-172">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>

    - <span data-ttu-id="ff084-173">Excel Online : [Chargement de versions test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="ff084-173">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>

    - <span data-ttu-id="ff084-174">iPad et Mac : [Chargement de versions test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="ff084-174">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

5. <span data-ttu-id="ff084-175">Dans le menu **Accueil**, sélectionnez **Afficher le volet Office**.</span><span class="sxs-lookup"><span data-stu-id="ff084-175">On the **Home** menu, choose **Show Taskpane**.</span></span>

6. <span data-ttu-id="ff084-176">Dans le volet Office, sélectionnez **Créer un tableau**.</span><span class="sxs-lookup"><span data-stu-id="ff084-176">In the task pane, choose **Create Table**.</span></span>

    ![Didacticiel Excel -Créer un tableau](../images/excel-tutorial-create-table.png)

## <a name="filter-and-sort-a-table"></a><span data-ttu-id="ff084-178">Filtrer et trier un tableau</span><span class="sxs-lookup"><span data-stu-id="ff084-178">Filter and sort a table</span></span>

<span data-ttu-id="ff084-179">Dans cette étape du didacticiel, vous allez filtrer et trier le tableau que vous avez créé précédemment.</span><span class="sxs-lookup"><span data-stu-id="ff084-179">In this step of the tutorial, you'll filter and sort the table that you created previously.</span></span>

### <a name="filter-the-table"></a><span data-ttu-id="ff084-180">Filtrage du tableau</span><span class="sxs-lookup"><span data-stu-id="ff084-180">Filter the table</span></span>

1. <span data-ttu-id="ff084-181">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="ff084-181">Open the project in your code editor.</span></span>

2. <span data-ttu-id="ff084-182">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="ff084-182">Open the file index.html.</span></span>

3. <span data-ttu-id="ff084-183">Juste en dessous de la balise `div` qui contient le bouton `create-table`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="ff084-183">Just below the `div` that contains the `create-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="filter-table">Filter Table</button>
    </div>
    ```

4. <span data-ttu-id="ff084-184">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="ff084-184">Open the app.js file.</span></span>

5. <span data-ttu-id="ff084-185">Juste en dessous de la ligne qui attribue un gestionnaire de clic au bouton `create-table`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="ff084-185">Just below the line that assigns a click handler to the `create-table` button, add the following code:</span></span>

    ```js
    $('#filter-table').click(filterTable);
    ```

6. <span data-ttu-id="ff084-186">Ajoutez la fonction suivante juste après la fonction `createTable`.</span><span class="sxs-lookup"><span data-stu-id="ff084-186">Just below the `createTable` function, add the following function:</span></span>

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

7. <span data-ttu-id="ff084-p113">Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="ff084-p113">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="ff084-189">Le code obtient tout d’abord une référence à la colonne à filtrer en transférant le nom de la colonne à la méthode `getItem`, au lieu de transmettre son index à la méthode `getItemAt` comme le fait la méthode `createTable`.</span><span class="sxs-lookup"><span data-stu-id="ff084-189">The code first gets a reference to the column that needs filtering by passing the column name to the `getItem` method, instead of passing its index to the `getItemAt` method as the `createTable` method does.</span></span> <span data-ttu-id="ff084-190">Puisque les utilisateurs peuvent déplacer des colonnes de tableau, la colonne d’un index donné peut être modifiée après la création du tableau.</span><span class="sxs-lookup"><span data-stu-id="ff084-190">Since users can move table columns, the column at a given index might change after the table is created.</span></span> <span data-ttu-id="ff084-191">Par conséquent, il est préférable d’utiliser le nom de la colonne pour obtenir une référence de la colonne.</span><span class="sxs-lookup"><span data-stu-id="ff084-191">Hence, it is safer to use the column name to get a reference to the column.</span></span> <span data-ttu-id="ff084-192">Dans le didacticiel précédent, nous avons utilisé la méthode `getItemAt` en toute sécurité, car nous l’avons utilisée dans la même méthode que celle qui crée le tableau, il n’y a donc aucune chance qu’un utilisateur ait déplacé la colonne.</span><span class="sxs-lookup"><span data-stu-id="ff084-192">We used `getItemAt` safely in the preceding tutorial, because we used it in the very same method that creates the table, so there is no chance that a user has moved the column.</span></span>

   - <span data-ttu-id="ff084-193">La méthode `applyValuesFilter` est l’une des nombreuses méthodes de filtrage sur l’objet `Filter`.</span><span class="sxs-lookup"><span data-stu-id="ff084-193">The `applyValuesFilter` method is one of several filtering methods on the `Filter` object.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);
    ``` 

### <a name="sort-the-table"></a><span data-ttu-id="ff084-194">Tri du tableau</span><span class="sxs-lookup"><span data-stu-id="ff084-194">Sort the table</span></span>

1. <span data-ttu-id="ff084-195">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="ff084-195">Open the file index.html.</span></span>

2. <span data-ttu-id="ff084-196">En dessous de la balise `div` qui contient le bouton `filter-table`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="ff084-196">Below the `div` that contains the `filter-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="sort-table">Sort Table</button>
    </div>
    ```

3. <span data-ttu-id="ff084-197">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="ff084-197">Open the app.js file.</span></span>

4. <span data-ttu-id="ff084-198">Sous la ligne qui attribue un gestionnaire de clics au bouton `filter-table`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="ff084-198">Below the line that assigns a click handler to the `filter-table` button, add the following code:</span></span>

    ```js
    $('#sort-table').click(sortTable);
    ```

5. <span data-ttu-id="ff084-199">Ajoutez la fonction suivante après la fonction `filterTable`.</span><span class="sxs-lookup"><span data-stu-id="ff084-199">Below the `filterTable` function add the following function.</span></span>

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

6. <span data-ttu-id="ff084-p115">Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="ff084-p115">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="ff084-202">Le code crée un tableau d’objets `SortField` qui ne comporte qu’un seul membre puisque le complément ne trie que la colonne Merchant.</span><span class="sxs-lookup"><span data-stu-id="ff084-202">The code creates an array of `SortField` objects which has just one member since the add-in only sorts on the Merchant column.</span></span>

   - <span data-ttu-id="ff084-203">La propriété `key` d’un objet `SortField` est l’index de la colonne à trier qui part de zéro.</span><span class="sxs-lookup"><span data-stu-id="ff084-203">The `key` property of a `SortField` object is the zero-based index of the column to sort-on.</span></span>

   - <span data-ttu-id="ff084-204">Le membre `sort` d’un objet `Table` est un objet `TableSort`, et non une méthode.</span><span class="sxs-lookup"><span data-stu-id="ff084-204">The `sort` member of a `Table` is a `TableSort` object, not a method.</span></span> <span data-ttu-id="ff084-205">Les objets `SortField` sont transmis à la méthode `apply` de l’objet `TableSort`.</span><span class="sxs-lookup"><span data-stu-id="ff084-205">The `SortField`s are passed to the `TableSort` object's `apply` method.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var sortFields = [
        {
            key: 1,            // Merchant column
            ascending: false,
        }
    ];

    expensesTable.sort.apply(sortFields);
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="ff084-206">Test du complément</span><span class="sxs-lookup"><span data-stu-id="ff084-206">Test the add-in</span></span>

1. <span data-ttu-id="ff084-207">Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur \*\*Ctrl+C \*\*à deux reprises pour arrêter le serveur web en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="ff084-207">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="ff084-208">Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.</span><span class="sxs-lookup"><span data-stu-id="ff084-208">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="ff084-209">Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet.</span><span class="sxs-lookup"><span data-stu-id="ff084-209">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="ff084-210">Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build.</span><span class="sxs-lookup"><span data-stu-id="ff084-210">In order to do this, you need to kill the server process so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="ff084-211">Une fois la commande build exécutée, redémarrez le serveur.</span><span class="sxs-lookup"><span data-stu-id="ff084-211">After the build, you restart the server.</span></span> <span data-ttu-id="ff084-212">Les prochaines étapes vous permettent d’effectuer ce processus.</span><span class="sxs-lookup"><span data-stu-id="ff084-212">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="ff084-213">Exécutez la commande `npm run build` pour transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par Internet Explorer (qui est utilisé en arrière-plan par Excel pour exécuter les compléments Excel).</span><span class="sxs-lookup"><span data-stu-id="ff084-213">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="ff084-214">Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.</span><span class="sxs-lookup"><span data-stu-id="ff084-214">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="ff084-215">Rechargez le volet des tâches en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet des tâches** pour rouvrir le complément.</span><span class="sxs-lookup"><span data-stu-id="ff084-215">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="ff084-216">Si, pour une raison quelconque, le tableau ne se trouve pas dans la feuille de calcul ouverte, dans le volet Office, sélectionnez **Créer un tableau**.</span><span class="sxs-lookup"><span data-stu-id="ff084-216">If for any reason the table is not in the open worksheet, in the task pane, choose **Create Table**.</span></span>

6. <span data-ttu-id="ff084-217">Choisissez les boutons **Filtrer le tableau** et **Trier le tableau** dans n’importe quel ordre.</span><span class="sxs-lookup"><span data-stu-id="ff084-217">Choose the **Filter Table** and **Sort Table** buttons, in either order.</span></span>

    ![Didacticiel Excel- Filtrer et trier un tableau](../images/excel-tutorial-filter-and-sort-table.png)

## <a name="create-a-chart"></a><span data-ttu-id="ff084-219">Création d’un graphique (chart)</span><span class="sxs-lookup"><span data-stu-id="ff084-219">Create a chart</span></span>

<span data-ttu-id="ff084-220">Dans cette étape du didacticiel, vous créerez un graphique à l’aide de données provenant du tableau précédemment créé, puis vous mettrez en forme le graphique.</span><span class="sxs-lookup"><span data-stu-id="ff084-220">In this step of the tutorial, you'll create a chart using data from the table that you created previously, and then format the chart.</span></span>

### <a name="chart-a-chart-using-table-data"></a><span data-ttu-id="ff084-221">Un graphique à l’aide de données du tableau de graphique (chart)</span><span class="sxs-lookup"><span data-stu-id="ff084-221">Chart a chart using table data</span></span>

1. <span data-ttu-id="ff084-222">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="ff084-222">Open the project in your code editor.</span></span>

2. <span data-ttu-id="ff084-223">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="ff084-223">Open the file index.html.</span></span>

3. <span data-ttu-id="ff084-224">En dessous de la balise `div` qui contient le bouton `sort-table`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="ff084-224">Below the `div` that contains the `sort-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="create-chart">Create Chart</button>
    </div>
    ```

4. <span data-ttu-id="ff084-225">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="ff084-225">Open the app.js file.</span></span>

5. <span data-ttu-id="ff084-226">Sous la ligne qui attribue un gestionnaire de clics au bouton `sort-chart`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="ff084-226">Below the line that assigns a click handler to the `sort-chart` button, add the following code:</span></span>

    ```js
    $('#create-chart').click(createChart);
    ```

6. <span data-ttu-id="ff084-227">Sous la fonction `sortTable`, ajoutez la fonction suivante.</span><span class="sxs-lookup"><span data-stu-id="ff084-227">Below the `sortTable` function add the following function.</span></span>

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

7. <span data-ttu-id="ff084-p119">Remplacez `TODO1` par le code suivant. Pour exclure la ligne d’en-tête, le code utilise la méthode `Table.getDataBodyRange` pour obtenir la plage de données à représenter sous forme de graphique à la place de la méthode `getRange`.</span><span class="sxs-lookup"><span data-stu-id="ff084-p119">Replace `TODO1` with the following code. Note that in order to exclude the header row, the code uses the `Table.getDataBodyRange` method to get the range of data you want to chart instead of the `getRange` method.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var dataRange = expensesTable.getDataBodyRange();
    ```

8. <span data-ttu-id="ff084-230">Remplacez `TODO2` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="ff084-230">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="ff084-231">Notez les paramètres suivants :</span><span class="sxs-lookup"><span data-stu-id="ff084-231">Note the following parameters:</span></span>

   - <span data-ttu-id="ff084-p121">Le premier paramètre transmis à la méthode `add` spécifie le type de graphique. Il en existe plusieurs dizaines de types.</span><span class="sxs-lookup"><span data-stu-id="ff084-p121">The first parameter to the `add` method specifies the type of chart. There are several dozen types.</span></span>

   - <span data-ttu-id="ff084-234">Le deuxième paramètre spécifie la plage de données à inclure dans le graphique.</span><span class="sxs-lookup"><span data-stu-id="ff084-234">The second parameter specifies the range of data to include in the chart.</span></span>

   - <span data-ttu-id="ff084-235">Le troisième paramètre détermine si une série de points de données provenant du tableau doit être représentée sous forme de graphique par ligne ou par colonne.</span><span class="sxs-lookup"><span data-stu-id="ff084-235">The third parameter determines whether a series of data points from the table should be charted row-wise or column-wise.</span></span> <span data-ttu-id="ff084-236">L’option `auto` demande à Excel de déterminer la meilleure méthode.</span><span class="sxs-lookup"><span data-stu-id="ff084-236">The option `auto` tells Excel to decide the best method.</span></span>

    ```js
    var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ```

9. <span data-ttu-id="ff084-237">Remplacez `TODO3` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="ff084-237">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="ff084-238">La majeure partie du code est explicite.</span><span class="sxs-lookup"><span data-stu-id="ff084-238">Most of this code is self-explanatory.</span></span> <span data-ttu-id="ff084-239">Remarque :</span><span class="sxs-lookup"><span data-stu-id="ff084-239">Note:</span></span>
   
   - <span data-ttu-id="ff084-240">Les paramètres de la méthode `setPosition` spécifient les cellules situées en haut à gauche et en bas à droite de la zone de feuille de calcul devant contenir le graphique.</span><span class="sxs-lookup"><span data-stu-id="ff084-240">The parameters to the `setPosition` method specify the upper left and lower right cells of the worksheet area that should contain the chart.</span></span> <span data-ttu-id="ff084-241">Excel peut ajuster des éléments, tels que la largeur de ligne pour que le graphique s’affiche correctement dans l’espace attribué.</span><span class="sxs-lookup"><span data-stu-id="ff084-241">Excel can adjust things like line width to make the chart look good in the space it has been given.</span></span>
   
   - <span data-ttu-id="ff084-242">Une « série » est un ensemble de points de données dans une colonne du tableau.</span><span class="sxs-lookup"><span data-stu-id="ff084-242">A "series" is a set of data points from a column of the table.</span></span> <span data-ttu-id="ff084-243">Étant donné qu’il n’existe qu’une seule colonne autre que de type chaîne dans le tableau, Excel déduit que la colonne est la seule colonne de points de données pour le graphique.</span><span class="sxs-lookup"><span data-stu-id="ff084-243">Since there is only one non-string column in the table, Excel infers that the column is the only column of data points to chart.</span></span> <span data-ttu-id="ff084-244">Il interprète les autres colonnes comme des étiquettes de graphique.</span><span class="sxs-lookup"><span data-stu-id="ff084-244">It interprets the other columns as chart labels.</span></span> <span data-ttu-id="ff084-245">Par conséquent, il y aura simplement une série dans le graphique et un index 0.</span><span class="sxs-lookup"><span data-stu-id="ff084-245">So there will be just one series in the chart and it will have index 0.</span></span> <span data-ttu-id="ff084-246">Il s’agit de celle à étiqueter avec « Valeur en € ».</span><span class="sxs-lookup"><span data-stu-id="ff084-246">This is the one to label with "Value in €".</span></span>

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in €';
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="ff084-247">Test du complément</span><span class="sxs-lookup"><span data-stu-id="ff084-247">Test the add-in</span></span>

1. <span data-ttu-id="ff084-248">Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur \*\*Ctrl+C \*\*à deux reprises pour arrêter le serveur web en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="ff084-248">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="ff084-249">Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.</span><span class="sxs-lookup"><span data-stu-id="ff084-249">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="ff084-250">Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet.</span><span class="sxs-lookup"><span data-stu-id="ff084-250">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="ff084-251">Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build.</span><span class="sxs-lookup"><span data-stu-id="ff084-251">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="ff084-252">Une fois la commande build exécutée, redémarrez le serveur.</span><span class="sxs-lookup"><span data-stu-id="ff084-252">After the build, you restart the server.</span></span> <span data-ttu-id="ff084-253">Les prochaines étapes vous permettent d’effectuer ce processus.</span><span class="sxs-lookup"><span data-stu-id="ff084-253">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="ff084-254">Exécutez la commande `npm run build` pour transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par Internet Explorer (qui est utilisé en arrière-plan par Excel pour exécuter les compléments Excel).</span><span class="sxs-lookup"><span data-stu-id="ff084-254">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="ff084-255">Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.</span><span class="sxs-lookup"><span data-stu-id="ff084-255">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="ff084-256">Rechargez le volet des tâches en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet des tâches** pour rouvrir le complément.</span><span class="sxs-lookup"><span data-stu-id="ff084-256">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="ff084-257">Si pour une raison quelconque le tableau n’est pas dans la feuille de calcul ouverte, dans le volet Office, sélectionnez **Créer un tableau**, puis les boutons **Filtrer le tableau** et **Trier le tableau** dans n’importe quel ordre.</span><span class="sxs-lookup"><span data-stu-id="ff084-257">If for any reason the table is not in the open worksheet, in the task pane, choose **Create Table** and then **Filter Table** and **Sort Table** buttons, in either order.</span></span>

6. <span data-ttu-id="ff084-258">Sélectionnez le bouton **Créer un graphique**.</span><span class="sxs-lookup"><span data-stu-id="ff084-258">Choose the **Create Chart** button.</span></span> <span data-ttu-id="ff084-259">Un graphique est créé dans lequel seules les données provenant des lignes filtrées sont incluses.</span><span class="sxs-lookup"><span data-stu-id="ff084-259">A chart is created and only the data from the rows that have been filtered are included.</span></span> <span data-ttu-id="ff084-260">Les étiquettes sur les points de données en bas sont organisées selon l’ordre de tri du graphique, à savoir les noms de marchand par ordre alphabétique inversé.</span><span class="sxs-lookup"><span data-stu-id="ff084-260">The labels on the data points across the bottom are in the sort order of the chart; that is, merchant names in reverse alphabetical order.</span></span>

    ![Didacticiel Excel -Créer un graphique (chart)](../images/excel-tutorial-create-chart.png)

## <a name="freeze-a-table-header"></a><span data-ttu-id="ff084-262">Figer un en-tête de tableau</span><span class="sxs-lookup"><span data-stu-id="ff084-262">Freeze a table header in place</span></span>

<span data-ttu-id="ff084-263">Lorsqu’un tableau est tellement long que l’utilisateur doit le faire défiler pour afficher les lignes suivantes, la ligne d’en-tête peut être masquée.</span><span class="sxs-lookup"><span data-stu-id="ff084-263">When a table is long enough that a user must scroll to see some rows, the header row can scroll out of sight.</span></span> <span data-ttu-id="ff084-264">Dans cette étape du didacticiel, vous allez figer la ligne d’en-tête du tableau que vous avez créé précédemment, afin qu’elle reste visible même lorsque l’utilisateur fait défiler la feuille de calcul vers le bas.</span><span class="sxs-lookup"><span data-stu-id="ff084-264">In this step of the tutorial, you'll freeze the header row of the table that you created previously, so that it remains visible even as the user scrolls down the worksheet.</span></span>

### <a name="freeze-the-tables-header-row"></a><span data-ttu-id="ff084-265">Figer la ligne d’en-tête du tableau</span><span class="sxs-lookup"><span data-stu-id="ff084-265">Freeze the table's header row</span></span>

1. <span data-ttu-id="ff084-266">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="ff084-266">Open the project in your code editor.</span></span>

2. <span data-ttu-id="ff084-267">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="ff084-267">Open the file index.html.</span></span>

3. <span data-ttu-id="ff084-268">En dessous de la balise `div` qui contient le bouton `create-chart`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="ff084-268">Below the `div` that contains the `create-chart` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="freeze-header">Freeze Header</button>
    </div>
    ```

4. <span data-ttu-id="ff084-269">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="ff084-269">Open the app.js file.</span></span>

5. <span data-ttu-id="ff084-270">En dessous de la ligne qui attribue un gestionnaire de clic au bouton `create-chart`, ajoutez le code suivant :</span><span class="sxs-lookup"><span data-stu-id="ff084-270">Below the line that assigns a click handler to the `create-chart` button, add the following code:</span></span>

    ```js
    $('#freeze-header').click(freezeHeader);
    ```

6. <span data-ttu-id="ff084-271">En dessous de la fonction `createChart`, ajoutez la fonction suivante :</span><span class="sxs-lookup"><span data-stu-id="ff084-271">Below the `createChart` function add the following function:</span></span>

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

7. <span data-ttu-id="ff084-p130">Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="ff084-p130">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="ff084-274">La collection `Worksheet.freezePanes` est un ensemble de volets de la feuille de calcul qui sont épinglés, c’est-à-dire figés, lorsque vous faites défiler la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="ff084-274">The `Worksheet.freezePanes` collection is a set of panes in the worksheet that are pinned, or frozen, in place when the worksheet is scrolled.</span></span>

   - <span data-ttu-id="ff084-p131">La méthode `freezeRows` prend comme paramètre le nombre de lignes, à partir du haut, qui doivent être figées. Nous transmettons `1` pour figer la première ligne.</span><span class="sxs-lookup"><span data-stu-id="ff084-p131">The `freezeRows` method takes as a parameter the number of rows, from the top that are to be pinned in place. We pass `1` to pin the first row in place.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="ff084-277">Test du complément</span><span class="sxs-lookup"><span data-stu-id="ff084-277">Test the add-in</span></span>

1. <span data-ttu-id="ff084-278">Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur \*\*Ctrl+C \*\*à deux reprises pour arrêter le serveur web en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="ff084-278">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="ff084-279">Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.</span><span class="sxs-lookup"><span data-stu-id="ff084-279">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="ff084-280">Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet.</span><span class="sxs-lookup"><span data-stu-id="ff084-280">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="ff084-281">Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build.</span><span class="sxs-lookup"><span data-stu-id="ff084-281">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="ff084-282">Une fois la commande build exécutée, redémarrez le serveur.</span><span class="sxs-lookup"><span data-stu-id="ff084-282">After the build, you restart the server.</span></span> <span data-ttu-id="ff084-283">Les prochaines étapes vous permettent d’effectuer ce processus.</span><span class="sxs-lookup"><span data-stu-id="ff084-283">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="ff084-284">Exécutez la commande `npm run build` pour transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par Internet Explorer (qui est utilisé en arrière-plan par Excel pour exécuter les compléments Excel).</span><span class="sxs-lookup"><span data-stu-id="ff084-284">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="ff084-285">Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.</span><span class="sxs-lookup"><span data-stu-id="ff084-285">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="ff084-286">Rechargez le volet Office en le fermant, puis dans le menu **Accueil**, sélectionnez **Afficher le volet Office** pour rouvrir le complément.</span><span class="sxs-lookup"><span data-stu-id="ff084-286">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="ff084-287">Si le tableau est dans la feuille de calcul, supprimez-le.</span><span class="sxs-lookup"><span data-stu-id="ff084-287">If the table is in the worksheet, delete it.</span></span>

6. <span data-ttu-id="ff084-288">Dans le volet Office, sélectionnez **Créer un tableau**.</span><span class="sxs-lookup"><span data-stu-id="ff084-288">In the task pane, choose **Create Table**.</span></span>

7. <span data-ttu-id="ff084-289">Sélectionnez le bouton **Freeze Header**.</span><span class="sxs-lookup"><span data-stu-id="ff084-289">Choose the **Freeze Header** button.</span></span>

8. <span data-ttu-id="ff084-290">Faites suffisamment défiler la feuille de calcul vers le bas pour voir que l’en-tête du tableau est toujours visible dans la partie supérieure même lorsque les lignes du haut sont masquées.</span><span class="sxs-lookup"><span data-stu-id="ff084-290">Scroll down the worksheet enough to to see that the table header remains visible at the top even when the higher rows scroll out of sight.</span></span>

    ![Didacticiel Excel-Figer l’en-tête](../images/excel-tutorial-freeze-header.png)

## <a name="protect-a-worksheet"></a><span data-ttu-id="ff084-292">Protéger une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="ff084-292">Protect a worksheet from changes</span></span>

<span data-ttu-id="ff084-293">Dans cette étape du didacticiel, vous allez ajouter un autre bouton au ruban qui, lorsque l’utilisateur clique dessus, exécute une fonction qui vous allez définir et qui active/désactive la protection de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="ff084-293">In this step of the tutorial, you'll add another button to the ribbon that, when chosen, executes a function that you'll define to toggle worksheet protection on and off.</span></span>

### <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a><span data-ttu-id="ff084-294">Configuration du manifeste pour ajouter un deuxième bouton de ruban</span><span class="sxs-lookup"><span data-stu-id="ff084-294">Configure the manifest to add a second ribbon button</span></span>

1. <span data-ttu-id="ff084-295">Ouvrez le fichier manifeste my-office-add-in-manifest.xml.</span><span class="sxs-lookup"><span data-stu-id="ff084-295">Open the manifest file my-office-add-in-manifest.xml.</span></span>

2. <span data-ttu-id="ff084-296">Recherchez l’élément `<Control>`.</span><span class="sxs-lookup"><span data-stu-id="ff084-296">Find the `<Control>` element.</span></span> <span data-ttu-id="ff084-297">Cet élément définit le bouton **Afficher le volet des pages** sur le ruban **Accueil** que vous utilisez pour lancer le complément.</span><span class="sxs-lookup"><span data-stu-id="ff084-297">This element defines the **Show Taskpane** button on the **Home** ribbon you have been using to launch the add-in.</span></span> <span data-ttu-id="ff084-298">Nous allons ajouter un deuxième bouton au même groupe sur le ruban **Accueil**.</span><span class="sxs-lookup"><span data-stu-id="ff084-298">We're going to add a second button to the same group on the **Home** ribbon.</span></span> <span data-ttu-id="ff084-299">Entre la balise Control de fin (`</Control>`) et la balise Group de fin (`</Group>`), ajoutez le balisage suivant.</span><span class="sxs-lookup"><span data-stu-id="ff084-299">In between the end Control tag (`</Control>`) and the end Group tag (`</Group>`), add the following markup.</span></span>

    ```xml
    <Control xsi:type="Button" id="<!--TODO1: Unique (in manifest) name for button -->">
        <Label resid="<!--TODO2: Button label -->" />
        <Supertip>            
            <Title resid="<!-- TODO3: Button tool tip title -->" />
            <Description resid="<!-- TODO4: Button tool tip description -->" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="<!-- TODO5: Specify the type of action-->">
            <!-- TODO6: Identify the function.-->
        </Action>
    </Control>
    ```

3. <span data-ttu-id="ff084-300">Remplacez `TODO1` par une chaîne qui attribue un ID unique au bouton au sein de ce fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="ff084-300">Replace `TODO1` with a string that gives the button an ID that is unique within this manifest file.</span></span> <span data-ttu-id="ff084-301">Étant donné que notre bouton va activer ou désactiver la protection de la feuille de calcul, utilisez « ToggleProtection ».</span><span class="sxs-lookup"><span data-stu-id="ff084-301">Since our button is going to toggle protection of the worksheet on and off, use "ToggleProtection".</span></span> <span data-ttu-id="ff084-302">Lorsque vous avez terminé, la balise Control de début doit ressembler à ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="ff084-302">When you are done, the entire start Control tag should look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. <span data-ttu-id="ff084-303">Les trois éléments `TODO` suivants définissent les éléments « resid », c’est-à-dire les ID de ressource.</span><span class="sxs-lookup"><span data-stu-id="ff084-303">The next three `TODO`s set "resid"s, which is short for resource ID.</span></span> <span data-ttu-id="ff084-304">Une ressource est une chaîne. Vous allez créer ces trois chaînes lors d’une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="ff084-304">A resource is a string, and you'll create these three strings in a later step.</span></span> <span data-ttu-id="ff084-305">Pour l’instant, vous devez attribuer des ID aux ressources.</span><span class="sxs-lookup"><span data-stu-id="ff084-305">For now, you need to give IDs to the resources.</span></span> <span data-ttu-id="ff084-306">L’étiquette du bouton doit indiquer « Toggle Protection », mais l’*ID* de cette chaîne doit être « ProtectionButtonLabel », donc l’élément `Label` terminé doit ressembler au code suivant :</span><span class="sxs-lookup"><span data-stu-id="ff084-306">The button label should read "Toggle Protection", but the *ID* of this string should be "ProtectionButtonLabel", so the completed `Label` element should look like the following code:</span></span>

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. <span data-ttu-id="ff084-307">L’élément `SuperTip` définit l’info-bulle du bouton.</span><span class="sxs-lookup"><span data-stu-id="ff084-307">The `SuperTip` element defines the tool tip for the button.</span></span> <span data-ttu-id="ff084-308">Le titre de l’info-bulle doit être identique à l’étiquette du bouton, nous utilisons donc le même ID de ressource : « ProtectionButtonLabel ».</span><span class="sxs-lookup"><span data-stu-id="ff084-308">The tool tip title should be the same as the button label, so we use the very same resource ID: "ProtectionButtonLabel".</span></span> <span data-ttu-id="ff084-309">La description de l’info-bulle sera « Cliquez pour activer/désactiver la protection de la feuille de calcul ».</span><span class="sxs-lookup"><span data-stu-id="ff084-309">The tool tip description will be "Click to turn protection of the worksheet on and off".</span></span> <span data-ttu-id="ff084-310">Néanmoins, l’élément `ID` doit être « ProtectionButtonToolTip ».</span><span class="sxs-lookup"><span data-stu-id="ff084-310">But the `ID` should be "ProtectionButtonToolTip".</span></span> <span data-ttu-id="ff084-311">Ainsi, lorsque vous avez terminé, l’ensemble du balisage `SuperTip` doit ressembler au code suivant :</span><span class="sxs-lookup"><span data-stu-id="ff084-311">So, when you are done, the whole `SuperTip` markup should look like the following code:</span></span> 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > <span data-ttu-id="ff084-312">Dans un complément de production, vous n’utiliseriez pas la même icône pour deux boutons différents, mais pour simplifier ce didacticiel, nous allons le faire.</span><span class="sxs-lookup"><span data-stu-id="ff084-312">In a production add-in, you would not want to use the same icon for two different buttons; but to simplify this tutorial, we'll do that.</span></span> <span data-ttu-id="ff084-313">Par conséquent, le balisage `Icon` de notre nouvel élément `Control` est simplement une copie de l’élément `Icon` provenant de l’élément `Control` existant.</span><span class="sxs-lookup"><span data-stu-id="ff084-313">So the `Icon` markup in our new `Control` is just a copy of the `Icon` element from the existing `Control`.</span></span> 

6. <span data-ttu-id="ff084-314">Le type de l’élément `Action` se trouvant à l’intérieur de l’élément `Control` d’origine qui était déjà présent dans le fichier manifeste est défini sur `ShowTaskpane`, mais notre nouveau bouton ne va pas ouvrir un volet Office, il va exécuter une fonction personnalisée que vous allez créer à une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="ff084-314">The `Action` element inside the original `Control` element that was already present in the manifest, has its type set to `ShowTaskpane`, but our new button isn't going to open a task pane; it's going to run a custom function that you create in a later step.</span></span> <span data-ttu-id="ff084-315">Il faut donc remplacer `TODO5` par `ExecuteFunction`, c’est-à-dire le type d’action pour les boutons qui déclenchent des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="ff084-315">So replace `TODO5` with `ExecuteFunction` which is the action type for buttons that trigger custom functions.</span></span> <span data-ttu-id="ff084-316">La balise `Action` de début doit ressembler au code suivant :</span><span class="sxs-lookup"><span data-stu-id="ff084-316">The start `Action` tag should look like the following code:</span></span>
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. <span data-ttu-id="ff084-317">L’élément `Action` d’origine possède des éléments enfants qui spécifient un ID de volet Office ainsi qu’une URL de la page qui doit être ouverte dans le volet Office.</span><span class="sxs-lookup"><span data-stu-id="ff084-317">The original `Action` element has child elements that specify a task pane ID and a URL of the page that should be opened in the task pane.</span></span> <span data-ttu-id="ff084-318">Toutefois, un élément `Action` de type `ExecuteFunction` comporte un élément enfant unique qui nomme la fonction que le contrôle exécute.</span><span class="sxs-lookup"><span data-stu-id="ff084-318">But an `Action` element of the `ExecuteFunction` type has a single child element that names the function that the control executes.</span></span> <span data-ttu-id="ff084-319">Vous créerez cette fonction à une étape ultérieure, et la nommerez `toggleProtection`.</span><span class="sxs-lookup"><span data-stu-id="ff084-319">You'll create that function in a later step, and it will be called `toggleProtection`.</span></span> <span data-ttu-id="ff084-320">Par conséquent, remplacez `TODO6` par le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="ff084-320">So, replace `TODO6` with the following markup:</span></span>
 
    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    <span data-ttu-id="ff084-321">Le balisage `Control` complet doit à présent ressembler à ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="ff084-321">The entire `Control` markup should now look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
        <Label resid="ProtectionButtonLabel" />
        <Supertip>            
            <Title resid="ProtectionButtonLabel" />
            <Description resid="ProtectionButtonToolTip" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="ExecuteFunction">
           <FunctionName>toggleProtection</FunctionName>
        </Action>
    </Control>
    ```

8. <span data-ttu-id="ff084-322">Faites défiler vers le bas jusqu’à la section `Resources` du manifeste.</span><span class="sxs-lookup"><span data-stu-id="ff084-322">Scroll down to the `Resources` section of the manifest.</span></span>

9. <span data-ttu-id="ff084-323">Ajoutez le balisage suivant en tant qu’enfant de l’élément `bt:ShortStrings`.</span><span class="sxs-lookup"><span data-stu-id="ff084-323">Add the following markup as a child of the `bt:ShortStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. <span data-ttu-id="ff084-324">Ajoutez le balisage suivant en tant qu’enfant de l’élément `bt:LongStrings`.</span><span class="sxs-lookup"><span data-stu-id="ff084-324">Add the following markup as a child of the `bt:LongStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. <span data-ttu-id="ff084-325">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="ff084-325">Save the file.</span></span>

### <a name="create-the-function-that-protects-the-sheet"></a><span data-ttu-id="ff084-326">Création de la fonction qui protège la feuille</span><span class="sxs-lookup"><span data-stu-id="ff084-326">Create the function that protects the sheet</span></span>

1. <span data-ttu-id="ff084-327">Ouvrez le fichier \function-file\function-file.js.</span><span class="sxs-lookup"><span data-stu-id="ff084-327">Open the file \function-file\function-file.js.</span></span>

2. <span data-ttu-id="ff084-328">Le fichier possède déjà une expression de fonction appelée immédiatement (IIFE).</span><span class="sxs-lookup"><span data-stu-id="ff084-328">The file already has an Immediately Invoked Function Expression (IFFE).</span></span> <span data-ttu-id="ff084-329">Aucune logique d’initialisation personnalisée n’est nécessaire, donc laissez la fonction qui a été attribuée à `Office.initialize` avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="ff084-329">No custom initialization logic is needed, so leave the function that is assigned to `Office.initialize` with an empty body.</span></span> <span data-ttu-id="ff084-330">(Mais ne la supprimez pas.</span><span class="sxs-lookup"><span data-stu-id="ff084-330">(But do not delete it.</span></span> <span data-ttu-id="ff084-331">La propriété `Office.initialize` ne peut pas être null ou non définie.) *En dehors de l’IIFE*, ajoutez le code suivant.</span><span class="sxs-lookup"><span data-stu-id="ff084-331">The `Office.initialize` property cannot be null or undefined.) *Outside of the IIFE*, add the following code.</span></span> <span data-ttu-id="ff084-332">Notez que nous spécifions un paramètre `args` pour la méthode et que la toute dernière ligne de la méthode appelle `args.completed`.</span><span class="sxs-lookup"><span data-stu-id="ff084-332">Note that we specify an `args` parameter to the method and the very last line of the method calls `args.completed`.</span></span> <span data-ttu-id="ff084-333">Il s’agit d’une condition requise pour toutes les commandes de type **ExecuteFunction**.</span><span class="sxs-lookup"><span data-stu-id="ff084-333">This is a requirement for all add-in commands of type **ExecuteFunction**.</span></span> <span data-ttu-id="ff084-334">Elle signale à l’application hôte Office que la fonction est terminée et que l’interface utilisateur est à nouveau réactive.</span><span class="sxs-lookup"><span data-stu-id="ff084-334">It signals the Office host application that the function has finished and the UI can become responsive again.</span></span>

    ```js
    function toggleProtection(args) {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to reverse the protection status of the current worksheet.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```

3. <span data-ttu-id="ff084-335">Remplacez `TODO1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="ff084-335">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="ff084-336">Ce code utilise la propriété de protection de l’objet de feuille de calcul dans un modèle de bouton bascule standard.</span><span class="sxs-lookup"><span data-stu-id="ff084-336">This code uses the worksheet object's protection property in a standard toggle pattern.</span></span> <span data-ttu-id="ff084-337">L’élément `TODO2` sera expliqué dans la section suivante.</span><span class="sxs-lookup"><span data-stu-id="ff084-337">The `TODO2` will be explained in the next section.</span></span>

    ```js
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // TODO2: Queue command to load the sheet's "protection.protected" property from
    //        the document and re-synchronize the document and task pane.

     if (sheet.protection.protected) {
        sheet.protection.unprotect();
    } else {
        sheet.protection.protect();
    }
    ``` 

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="ff084-338">Ajoutez du code pour récupérer des propriétés de document dans les objets de script du volet Office</span><span class="sxs-lookup"><span data-stu-id="ff084-338">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="ff084-339">Dans toutes les fonctions précédentes de cette série de didacticiels, vous avez mis en file d’attente des commandes pour écrire (*write*) dans le document Office.</span><span class="sxs-lookup"><span data-stu-id="ff084-339">In all the earlier functions in this series of tutorials, you queued commands to *write* to the Office document.</span></span> <span data-ttu-id="ff084-340">Chaque fonction se terminait par un appel de la méthode `context.sync()` qui envoie les commandes en file d’attente au document pour qu’elles soient exécutées.</span><span class="sxs-lookup"><span data-stu-id="ff084-340">Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed.</span></span> <span data-ttu-id="ff084-341">Cependant, le code que vous avez ajouté dans la dernière étape appelle la propriété `sheet.protection.protected` et c’est une différence significative par rapport aux fonctions antérieures que vous avez écrites, car l’objet `sheet` est uniquement un objet de proxy qui existe dans le script de votre volet Office.</span><span class="sxs-lookup"><span data-stu-id="ff084-341">But the code you added in the last step calls the `sheet.protection.protected` property, and this is a significant difference from the earlier functions you wrote, because the `sheet` object is only a proxy object that exists in your task pane's script.</span></span> <span data-ttu-id="ff084-342">Il ne connaît pas l’état de protection réel du document, donc sa propriété `protection.protected` ne peut pas contenir une valeur réelle.</span><span class="sxs-lookup"><span data-stu-id="ff084-342">It doesn't know what the actual protection state of the document is, so its `protection.protected` property can't have a real value.</span></span> <span data-ttu-id="ff084-343">Tout d’abord, il faut récupérer l’état de protection dans le document et l’utiliser pour définir la valeur de `sheet.protection.protected`.</span><span class="sxs-lookup"><span data-stu-id="ff084-343">It is necessary to first fetch the protection status from the document and use it set the value of `sheet.protection.protected`.</span></span> <span data-ttu-id="ff084-344">Seulement ensuite, la propriété `sheet.protection.protected` peut être appelée sans générer d’exception.</span><span class="sxs-lookup"><span data-stu-id="ff084-344">Only then can `sheet.protection.protected` be called without causing an exception to be thrown.</span></span> <span data-ttu-id="ff084-345">Ce processus de récupération comporte trois étapes :</span><span class="sxs-lookup"><span data-stu-id="ff084-345">This fetching process has three steps:</span></span>

   1. <span data-ttu-id="ff084-346">Mettez en file d’attente une commande de chargement (c’est-à-dire, fetch) des propriétés que votre code doit lire.</span><span class="sxs-lookup"><span data-stu-id="ff084-346">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>

   2. <span data-ttu-id="ff084-347">Appelez la méthode `sync` de l’objet de contexte pour envoyer la commande mise en file d’attente vers le document pour exécution, et renvoyez les informations demandées.</span><span class="sxs-lookup"><span data-stu-id="ff084-347">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>

   3. <span data-ttu-id="ff084-348">Étant donné que la méthode `sync` est asynchrone, assurez-vous qu’elle est terminée avant que votre code appelle les propriétés qui ont été récupérées.</span><span class="sxs-lookup"><span data-stu-id="ff084-348">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="ff084-349">Ces étapes doivent être effectuées à chaque fois que votre code doit lire (*read*) des informations provenant du document Office.</span><span class="sxs-lookup"><span data-stu-id="ff084-349">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="ff084-p144">Dans la fonction `toggleProtection`, remplacez `TODO2` par le code suivant. Remarque :</span><span class="sxs-lookup"><span data-stu-id="ff084-p144">In the `toggleProtection` function, replace `TODO2` with the following code. Note:</span></span>
   
   - <span data-ttu-id="ff084-352">Chaque objet Excel possède une méthode `load`.</span><span class="sxs-lookup"><span data-stu-id="ff084-352">Every Excel object has a `load` method.</span></span> <span data-ttu-id="ff084-353">Vous spécifiez les propriétés de l’objet que vous voulez lire dans le paramètre en tant que chaîne de noms séparés par des virgules.</span><span class="sxs-lookup"><span data-stu-id="ff084-353">You specify the properties of the object that you want to read in the parameter as a string of comma-delimited names.</span></span> <span data-ttu-id="ff084-354">Dans ce cas, la propriété que vous devez lire est une sous-propriété de la propriété `protection`.</span><span class="sxs-lookup"><span data-stu-id="ff084-354">In this case, the property you need to read is a subproperty of the `protection` property.</span></span> <span data-ttu-id="ff084-355">Pour référence la sous-propriété, procédez presque exactement de la même façon que vous le feriez à n’importe quel autre emplacement de votre code, sauf que vous devez utiliser une barre oblique (« / ») au lieu d’un point « . ».</span><span class="sxs-lookup"><span data-stu-id="ff084-355">You reference the subproperty almost exactly as you would anywhere else in your code, with the exception that you use a forward slash ('/') character instead of a "." character.</span></span>

   - <span data-ttu-id="ff084-356">Pour être sûr que la logique de bouton bascule, qui lit `sheet.protection.protected`, ne s’exécute pas tant que la synchronisation (`sync`) n’est pas terminée et que l’élément `sheet.protection.protected` n’a pas été affecté à la valeur correcte récupérée à partir du document, elle sera déplacée (à l’étape suivante) dans une fonction `then` qui ne s’exécutera pas tant que la synchronisation (`sync`) ne sera pas terminée.</span><span class="sxs-lookup"><span data-stu-id="ff084-356">To ensure that the toggle logic, which reads `sheet.protection.protected`, does not run until after the `sync` is complete and the `sheet.protection.protected` has been assigned the correct value that is fetched from the document, it will be moved (in the next step) into a `then` function that won't run until the `sync` has completed.</span></span> 

    ```js
    sheet.load('protection/protected');
    return context.sync()
        .then(
            function() {
                // TODO3: Move the queued toggle logic here.
            }
        )
        // TODO4: Move the final call of `context.sync` here and ensure that it
        //        does not run until the toggle logic has been queued.
    ``` 

2. <span data-ttu-id="ff084-357">Il n’est pas possible que deux instructions `return` se trouvent dans le même chemin de code, donc supprimez la dernière ligne `return context.sync();` à la fin de la fonction `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="ff084-357">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Excel.run`.</span></span> <span data-ttu-id="ff084-358">Vous ajouterez un nouvel élément final `context.sync` dans une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="ff084-358">You will add a new final `context.sync`, in a later step.</span></span>

3. <span data-ttu-id="ff084-359">Coupez la structurer `if ... else` dans la fonction `toggleProtection` et collez-la à la place de `TODO3`.</span><span class="sxs-lookup"><span data-stu-id="ff084-359">Cut the `if ... else` structure in the `toggleProtection` function and paste it in place of `TODO3`.</span></span>

4. <span data-ttu-id="ff084-p147">Remplacez `TODO4` par le code suivant. Veuillez noter les informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="ff084-p147">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="ff084-362">Le fait de transmettre la méthode `sync` à une fonction `then` permet de s’assurer qu’elle n’est pas exécutée tant que `sheet.protection.unprotect()` ou `sheet.protection.protect()` n’a pas été mis en file d’attente.</span><span class="sxs-lookup"><span data-stu-id="ff084-362">Passing the `sync` method to a `then` function ensures that it does not run until either `sheet.protection.unprotect()` or `sheet.protection.protect()` has been queued.</span></span>

   - <span data-ttu-id="ff084-363">La méthode `then` appelle n’importe quelle fonction qui lui est transmise, et vous ne souhaitez pas appeler `sync` deux fois, donc omettez les parenthèses « () » à la fin de `context.sync`.</span><span class="sxs-lookup"><span data-stu-id="ff084-363">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of `context.sync`.</span></span>

    ```js
    .then(context.sync);
    ```

   <span data-ttu-id="ff084-364">Lorsque vous avez terminé, la fonction entière doit ressembler à ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="ff084-364">When you are done, the entire function should look like the following:</span></span>

    ```js
    function toggleProtection(args) {
        Excel.run(function (context) {            
          var sheet = context.workbook.worksheets.getActiveWorksheet();          
          sheet.load('protection/protected');

          return context.sync()
              .then(
                  function() {
                    if (sheet.protection.protected) {
                        sheet.protection.unprotect();
                    } else {
                        sheet.protection.protect();
                    }
                  }
              )
              .then(context.sync);
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```

### <a name="configure-the-script-loading-html-file"></a><span data-ttu-id="ff084-365">Configuration du fichier HTML de chargement de script</span><span class="sxs-lookup"><span data-stu-id="ff084-365">Configure the script-loading HTML file</span></span>

<span data-ttu-id="ff084-366">Ouvrez le fichier /function-file/function-file.html.</span><span class="sxs-lookup"><span data-stu-id="ff084-366">Open the /function-file/function-file.html file.</span></span> <span data-ttu-id="ff084-367">Il s’agit d’un fichier HTML sans interface utilisateur qui est appelé lorsque l’utilisateur appuie sur le bouton **Toggle Worksheet Protection**.</span><span class="sxs-lookup"><span data-stu-id="ff084-367">This is a UI-less HTML file that is called when the user presses the **Toggle Worksheet Protection** button.</span></span> <span data-ttu-id="ff084-368">Son objectif consiste à charger la méthode JavaScript qui doit s’exécuter lorsque l’utilisateur appuie sur le bouton.</span><span class="sxs-lookup"><span data-stu-id="ff084-368">Its purpose is to load the JavaScript method that should run when the button is pushed.</span></span> <span data-ttu-id="ff084-369">Vous n’allez pas modifier ce fichier.</span><span class="sxs-lookup"><span data-stu-id="ff084-369">You are not going to change this file.</span></span> <span data-ttu-id="ff084-370">Remarquez simplement que la deuxième balise `<script>` charge le fichier functionfile.js.</span><span class="sxs-lookup"><span data-stu-id="ff084-370">Simply note that the second `<script>` tag loads the functionfile.js.</span></span>

   > [!NOTE]
   > <span data-ttu-id="ff084-371">Le fichier function-file.html et le fichier function-file.js qu’il charge s’exécutent dans un processus Internet Explorer entièrement distinct dans le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="ff084-371">The function-file.html file and the function-file.js file that it loads run in an entirely separate IE process from the add-in's task pane.</span></span> <span data-ttu-id="ff084-372">Si le fichier function-file.js était transpilé dans le même fichier bundle.js en tant que fichier app.js, le complément devrait charger deux copies du fichier bundle.js, ce qui irait à l’encontre l’objectif de groupement.</span><span class="sxs-lookup"><span data-stu-id="ff084-372">If the function-file.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="ff084-373">En outre, le fichier function-file.js ne contient pas de code JavaScript car Internet Explorer ne prend pas en charge ce type de code.</span><span class="sxs-lookup"><span data-stu-id="ff084-373">In addition, the function-file.js file does not contain any JavaScript that is unsupported by IE.</span></span> <span data-ttu-id="ff084-374">C’est pour ces deux raisons que ce complément ne transpile pas le fichier function-file.js du tout.</span><span class="sxs-lookup"><span data-stu-id="ff084-374">For these two reasons, this add-in does not transpile the function-file.js at all.</span></span> 

### <a name="test-the-add-in"></a><span data-ttu-id="ff084-375">Test du complément</span><span class="sxs-lookup"><span data-stu-id="ff084-375">Test the add-in</span></span>

1. <span data-ttu-id="ff084-376">Fermez toutes les applications Office, y compris Excel.</span><span class="sxs-lookup"><span data-stu-id="ff084-376">Close all Office applications, including Excel.</span></span> 

2. <span data-ttu-id="ff084-377">Supprimez le cache Office en supprimant le contenu du dossier de cache.</span><span class="sxs-lookup"><span data-stu-id="ff084-377">Delete the Office cache by deleting the contents of the cache folder.</span></span> <span data-ttu-id="ff084-378">Cette opération est nécessaire pour effacer complètement de l’hôte l’ancienne version du complément.</span><span class="sxs-lookup"><span data-stu-id="ff084-378">This is necessary to completely clear the old version of the add-in from the host.</span></span> 

    - <span data-ttu-id="ff084-379">Pour Windows : `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="ff084-379">For Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

    - <span data-ttu-id="ff084-380">Pour Mac : `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="ff084-380">For Mac: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

3. <span data-ttu-id="ff084-381">Si, pour une quelconque raison, votre serveur n’est pas en cours d’exécution, accédez au dossier **Start** du projet et exécutez la commande `npm start` dans une fenêtre Git Bash ou une invite système Node.JS.</span><span class="sxs-lookup"><span data-stu-id="ff084-381">If for any reason, your server is not running, then in a Git Bash window, or Node.JS-enabled system prompt, navigate to the **Start** folder of the project and run the command `npm start`.</span></span> <span data-ttu-id="ff084-382">Vous n’avez pas besoin de recréer le projet, car le seul fichier JavaScript que vous avez modifié ne fait pas partie du fichier bundle.js créé.</span><span class="sxs-lookup"><span data-stu-id="ff084-382">You do not need to rebuild the project because the only JavaScript file you changed is not part of the built bundle.js.</span></span>

4. <span data-ttu-id="ff084-383">À l’aide de la nouvelle version du fichier manifeste modifié, répétez le processus de chargement de version test en utilisant l’une des méthodes suivantes.</span><span class="sxs-lookup"><span data-stu-id="ff084-383">Using the new version of the changed manifest file, repeat the sideloading process by using one of the following methods.</span></span> <span data-ttu-id="ff084-384">*Vous devez remplacer la copie précédente du fichier manifeste.*</span><span class="sxs-lookup"><span data-stu-id="ff084-384">*You should overwrite the previous copy of the manifest file.*</span></span>

    - <span data-ttu-id="ff084-385">Windows : [Chargement de version test des compléments Office](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="ff084-385">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>

    - <span data-ttu-id="ff084-386">Excel Online : [Chargement de versions test des compléments Office dans Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="ff084-386">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>

    - <span data-ttu-id="ff084-387">iPad et Mac : [Chargement de version test des compléments Office sur iPad et Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="ff084-387">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

5. <span data-ttu-id="ff084-388">Ouvrez une feuille de calcul dans Excel.</span><span class="sxs-lookup"><span data-stu-id="ff084-388">Open any worksheet in Excel.</span></span>

6. <span data-ttu-id="ff084-p153">Sur le ruban **Accueil**, sélectionnez **Toggle Worksheet Protection** (Activer/Désactiver la protection de la feuille de calcul). Notez que la plupart des contrôles figurant sur le ruban sont désactivés (et visuellement grisés) comme illustré dans la capture d’écran ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="ff084-p153">On the **Home** ribbon, choose **Toggle Worksheet Protection**. Note that most of the controls on the ribbon are disabled (and visually grayed-out) as seen in screenshot below.</span></span> 

7. <span data-ttu-id="ff084-391">Sélectionnez une cellule comme vous le feriez si vous vouliez modifier son contenu.</span><span class="sxs-lookup"><span data-stu-id="ff084-391">Choose a cell as you would if you wanted to change its content.</span></span> <span data-ttu-id="ff084-392">Vous rencontrez une erreur indiquant que la feuille de calcul est protégée.</span><span class="sxs-lookup"><span data-stu-id="ff084-392">You get an error telling you that the worksheet is protected.</span></span>

8. <span data-ttu-id="ff084-393">Sélectionnez **Toggle Worksheet Protection** à nouveau pour réactiver les contrôles. Vous pouvez alors modifier une nouvelle fois les valeurs de cellule.</span><span class="sxs-lookup"><span data-stu-id="ff084-393">Choose **Toggle Worksheet Protection** again, and the controls are reenabled, and you can change cell values again.</span></span>

    ![Didacticiel Excel-Ruban avec protection activée](../images/excel-tutorial-ribbon-with-protection-on.png)

## <a name="open-a-dialog"></a><span data-ttu-id="ff084-395">Ouvrir une boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="ff084-395">Open a dialog box</span></span>

<span data-ttu-id="ff084-396">Dans cette étape finale du didacticiel, vous allez ouvrir une boîte de dialogue dans votre complément, transmettre un message du processus de boîte de dialogue au processus de volet Office et fermer la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="ff084-396">In this final step of the tutorial, you'll open a dialog in your add-in, pass a message from the dialog process to the task pane process, and close the dialog.</span></span> <span data-ttu-id="ff084-397">Les boîtes de dialogue des compléments Office sont *non modales* : un utilisateur peut continuer à interagir à la fois avec le document dans l’application Office hôte et avec la page hôte dans le volet Office.</span><span class="sxs-lookup"><span data-stu-id="ff084-397">Office Add-in dialogs are *nonmodal*: a user can continue to interact with both the document in the host Office application and with the host page in the task pane.</span></span>

### <a name="create-the-dialog-page"></a><span data-ttu-id="ff084-398">Création de la page de boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="ff084-398">Create the dialog page</span></span>

1. <span data-ttu-id="ff084-399">Ouvrez le projet dans votre éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="ff084-399">Open the project in your code editor.</span></span>

2. <span data-ttu-id="ff084-400">Créez un fichier à la racine du projet (où se trouve le fichier index.html) et nommez-le popup.html.</span><span class="sxs-lookup"><span data-stu-id="ff084-400">Create a file in the root of the project (where index.html is) called popup.html.</span></span>

3. <span data-ttu-id="ff084-p156">Ajoutez le balisage suivant au fichier popup.html. Remarque :</span><span class="sxs-lookup"><span data-stu-id="ff084-p156">Add the following markup to popup.html. Note:</span></span>

   - <span data-ttu-id="ff084-403">La page comporte un champ `<input>`, dans lequel l’utilisateur entrera son nom, et un bouton qui permet d’envoyer le nom à la page dans le volet Office où il sera affiché.</span><span class="sxs-lookup"><span data-stu-id="ff084-403">The page has a `<input>` where the user will enter their name and a button that will send the name to the page in the task pane where it will be displayed.</span></span>

   - <span data-ttu-id="ff084-404">Le balisage charge un script appelé popup.js que vous allez créer dans une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="ff084-404">The markup loads a script called popup.js that you will create in a later step.</span></span>

   - <span data-ttu-id="ff084-405">Il charge également la bibliothèque Office.JS et jQuery, car ils seront utilisés dans popup.js.</span><span class="sxs-lookup"><span data-stu-id="ff084-405">It also loads the Office.JS library and jQuery because they will be used in popup.js.</span></span>

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

4. <span data-ttu-id="ff084-406">Créez un fichier à la racine du projet et nommez-le popup.js.</span><span class="sxs-lookup"><span data-stu-id="ff084-406">Create a file in the root of the project called popup.js.</span></span>

5. <span data-ttu-id="ff084-p157">Ajoutez le code suivant au fichier popup.js. Remarque :</span><span class="sxs-lookup"><span data-stu-id="ff084-p157">Add the following code to popup.js. Note:</span></span>

   - <span data-ttu-id="ff084-409">*Toutes les pages qui appellent des API dans la bibliothèque Office.JS doivent affecter une fonction à la propriété `Office.initialize`.*</span><span class="sxs-lookup"><span data-stu-id="ff084-409">*Every page that calls APIs in the Office.JS library must assign a function to the `Office.initialize` property.*</span></span> <span data-ttu-id="ff084-410">Si aucune initialisation n’est nécessaire, la fonction peut avoir un corps vide, mais la propriété ne doit pas être laissée indéfinie, affectée à null ni à une valeur qui n’est pas une fonction.</span><span class="sxs-lookup"><span data-stu-id="ff084-410">If no initialization is needed, then the function can have an empty body, but the property must not be left undefined, assigned to null or to a non-function value.</span></span> <span data-ttu-id="ff084-411">Pour voir un exemple, affichez le fichier app.js à la racine du projet.</span><span class="sxs-lookup"><span data-stu-id="ff084-411">For an example, see the app.js file in the project root.</span></span> <span data-ttu-id="ff084-412">Le code qui exécute l’affectation doit être exécuté avant tout appel à Office.JS ; l’affectation se trouve donc dans un fichier de script chargé par la page, comme dans ce cas.</span><span class="sxs-lookup"><span data-stu-id="ff084-412">The code that makes the assignment must run before any calls to Office.JS; hence the assignment is in a script file that is loaded by the page, as it is in this case.</span></span>
   
   - <span data-ttu-id="ff084-p159">La fonction `ready` jQuery est appelée à l’intérieur de la méthode `initialize`. Une règle quasi-universelle veut que le code de chargement, d’initialisation ou d’amorçage des autres bibliothèques JavaScript se trouve à l’intérieur de la fonction `Office.initialize`.</span><span class="sxs-lookup"><span data-stu-id="ff084-p159">The jQuery `ready` function is called inside the `initialize` method. It is an almost universal rule that the loading, initializing, or bootstrapping code of other JavaScript libraries should be inside the `Office.initialize` function.</span></span>

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

6. <span data-ttu-id="ff084-415">Remplacez `TODO1` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="ff084-415">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="ff084-416">Vous allez créer la fonction `sendStringToParentPage` à l’étape suivante.</span><span class="sxs-lookup"><span data-stu-id="ff084-416">You'll create the `sendStringToParentPage` function in the next step.</span></span>

    ```js
    $('#ok-button').click(sendStringToParentPage);
    ```

7. <span data-ttu-id="ff084-417">Remplacez `TODO2` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="ff084-417">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="ff084-418">La méthode `messageParent` transmet son paramètre à la page parent, qui est, dans ce cas, la page dans le volet Office.</span><span class="sxs-lookup"><span data-stu-id="ff084-418">The `messageParent` method passes its parameter to the parent page, in this case, the page in the task pane.</span></span> <span data-ttu-id="ff084-419">Le paramètre peut être une valeur booléenne ou une chaîne qui inclut tous les éléments qui peuvent être sérialisés en tant que chaîne, au format XML ou JSON.</span><span class="sxs-lookup"><span data-stu-id="ff084-419">The parameter can be a boolean or a string, which includes anything that can be serialized as a string, such as XML or JSON.</span></span>

    ```js
    function sendStringToParentPage() {
        var userName = $('#name-box').val();
        Office.context.ui.messageParent(userName);
    }
    ```

8. <span data-ttu-id="ff084-420">Enregistrez le fichier.</span><span class="sxs-lookup"><span data-stu-id="ff084-420">Save the file.</span></span>

   > [!NOTE]
   > <span data-ttu-id="ff084-421">Le fichier popup.html et le fichier popup.js qu’il charge s’exécutent dans un processus Internet Explorer entièrement séparé à partir du volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="ff084-421">The popup.html file, and the popup.js file that it loads, run in an entirely separate Internet Explorer process from the add-in's task pane.</span></span> <span data-ttu-id="ff084-422">Si le popup.js était transpilé dans le même fichier bundle.js en tant que fichier app.js, le complément devrait charger deux copies du fichier bundle.js, ce qui irait à l’encontre de l’objectif de groupement.</span><span class="sxs-lookup"><span data-stu-id="ff084-422">If the popup.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="ff084-423">En outre, le fichier popup.js ne contient pas de code JavaScript car Internet Explorer ne prend pas en charge ce type de code.</span><span class="sxs-lookup"><span data-stu-id="ff084-423">In addition, the popup.js file does not contain any JavaScript that is unsupported by IE.</span></span> <span data-ttu-id="ff084-424">C’est pour ces deux raisons que ce complément ne transpile pas le fichier popup.js du tout.</span><span class="sxs-lookup"><span data-stu-id="ff084-424">For these two reasons, this add-in does not transpile the popup.js file at all.</span></span>

### <a name="open-the-dialog-from-the-task-pane"></a><span data-ttu-id="ff084-425">Ouverture de la boîte de dialogue à partir du volet Office</span><span class="sxs-lookup"><span data-stu-id="ff084-425">Open the dialog from the task pane</span></span>

1. <span data-ttu-id="ff084-426">Ouvrez le fichier index.html.</span><span class="sxs-lookup"><span data-stu-id="ff084-426">Open the file index.html.</span></span>

2. <span data-ttu-id="ff084-427">Sous la balise `div` qui contient le bouton `freeze-header`, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="ff084-427">Below the `div` that contains the `freeze-header` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="open-dialog">Open Dialog</button>
    </div>
    ```

3. <span data-ttu-id="ff084-428">La boîte de dialogue invitera l’utilisateur à saisir son nom et transmettra ce nom au volet Office.</span><span class="sxs-lookup"><span data-stu-id="ff084-428">The dialog will prompt the user to enter a name and pass the user's name to the task pane.</span></span> <span data-ttu-id="ff084-429">Le volet Office s’affichera dans une étiquette.</span><span class="sxs-lookup"><span data-stu-id="ff084-429">The task pane will display it in a label.</span></span> <span data-ttu-id="ff084-430">Juste en dessous de la balise `div` que vous venez d’ajouter, ajoutez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="ff084-430">Immediately below the `div` that you just added, add the following markup:</span></span>

    ```html
    <div class="padding">
        <label id="user-name"></label>
    </div>
    ```

4. <span data-ttu-id="ff084-431">Ouvrez le fichier app.js.</span><span class="sxs-lookup"><span data-stu-id="ff084-431">Open the app.js file.</span></span>

5. <span data-ttu-id="ff084-432">Sous la ligne qui attribue un gestionnaire de clics au bouton `freeze-header`, ajoutez le code suivant.</span><span class="sxs-lookup"><span data-stu-id="ff084-432">Below the line that assigns a click handler to the `freeze-header` button, add the following code.</span></span> <span data-ttu-id="ff084-433">Vous allez créer la méthode `openDialog` à une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="ff084-433">You'll create the `openDialog` method in a later step.</span></span>

    ```js
    $('#open-dialog').click(openDialog);
    ```

6. <span data-ttu-id="ff084-p165">Ajoutez la déclaration suivante sous la fonction `freezeHeader`. Cette variable est utilisée pour conserver un objet dans le contexte d’exécution de la page parent qui agit en tant qu’intermédiaire pour le contexte d’exécution de la page de boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="ff084-p165">Below the `freezeHeader` function add the following declaration. This variable is used to hold an object in the parent page's execution context that acts as an intermediator to the dialog page's execution context.</span></span>

    ```js
    var dialog = null;
    ```

7. <span data-ttu-id="ff084-436">Sous la déclaration de la balise `dialog`, ajoutez la fonction suivante.</span><span class="sxs-lookup"><span data-stu-id="ff084-436">Below the declaration of `dialog`, add the following function.</span></span> <span data-ttu-id="ff084-437">Le plus important à remarquer à propos de ce code est ce qui ne s’y trouve *pas* : il n’y a aucun appel de `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="ff084-437">The important thing to notice about this code is what is *not* there: there is no call of `Excel.run`.</span></span> <span data-ttu-id="ff084-438">Cela est dû au fait que l’API d’ouverture de boîte de dialogue est partagée par tous les hôtes Office, elle fait donc partie de l’API commune JavaScript Office, pas de l’API spécifique d’Excel.</span><span class="sxs-lookup"><span data-stu-id="ff084-438">This is because the API to open a dialog is shared among all Office hosts, so it is part of the Office JavaScript Common API, not the Excel-specific API.</span></span>

    ```js
    function openDialog() {
        // TODO1: Call the Office Common API that opens a dialog
    }
    ```

8. <span data-ttu-id="ff084-p167">Remplacez `TODO1` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="ff084-p167">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="ff084-441">La méthode `displayDialogAsync` ouvre une boîte de dialogue au centre de l’écran.</span><span class="sxs-lookup"><span data-stu-id="ff084-441">The `displayDialogAsync` method opens a dialog in the center of the screen.</span></span>

   - <span data-ttu-id="ff084-442">Le premier paramètre est l’URL de la page à ouvrir.</span><span class="sxs-lookup"><span data-stu-id="ff084-442">The first parameter is the URL of the page to open.</span></span>

   - <span data-ttu-id="ff084-p168">Le deuxième paramètre transmet les options. `height` et `width` sont des pourcentages de la taille de la fenêtre de l’application Office.</span><span class="sxs-lookup"><span data-stu-id="ff084-p168">The second parameter passes options. `height` and `width` are percentages of the size of the Office application's window.</span></span>

    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},

        // TODO2: Add callback parameter.
    );
    ```

### <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a><span data-ttu-id="ff084-445">Traitement du message à partir de la boîte de dialogue et fermeture de la boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="ff084-445">Process the message from the dialog and close the dialog</span></span>

1. <span data-ttu-id="ff084-p169">Continuez dans le fichier app.js et remplacez `TODO2` par le code suivant. Remarque :</span><span class="sxs-lookup"><span data-stu-id="ff084-p169">Continue in the app.js file, and replace `TODO2` with the following code. Note:</span></span>

   - <span data-ttu-id="ff084-448">Le rappel est exécuté immédiatement après que la boîte de dialogue s’est ouverte correctement et avant que l’utilisateur ait pris une quelconque action dans la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="ff084-448">The callback is executed immediately after the dialog successfully opens and before the user has taken any action in the dialog.</span></span>

   - <span data-ttu-id="ff084-449">`result.value` représente l’objet qui agit comme un intermédiaire entre les contextes d’exécution des pages parent et de boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="ff084-449">The `result.value` is the object that acts as a kind of middleman between the execution contexts of the parent and dialog pages.</span></span>

   - <span data-ttu-id="ff084-450">La fonction `processMessage` sera créée à une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="ff084-450">The `processMessage` function will be created in a later step.</span></span> <span data-ttu-id="ff084-451">Ce gestionnaire traitera toutes les valeurs envoyées par la page de boîte de dialogue avec les appels de la fonction `messageParent`.</span><span class="sxs-lookup"><span data-stu-id="ff084-451">This handler will process any values that are sent from the dialog page with calls of the `messageParent` function.</span></span>

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. <span data-ttu-id="ff084-452">Sous la fonction `openDialog`, ajoutez la fonction suivante.</span><span class="sxs-lookup"><span data-stu-id="ff084-452">Below the `openDialog` function, add the following function.</span></span>

    ```js
    function processMessage(arg) {
        $('#user-name').text(arg.message);
        dialog.close();
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="ff084-453">Test du complément</span><span class="sxs-lookup"><span data-stu-id="ff084-453">Test the add-in</span></span>

1. <span data-ttu-id="ff084-454">Si la fenêtre Git Bash, ou l’invite système Node.JS, de l’étape précédente du didacticiel est encore ouverte, appuyez sur \*\*Ctrl+C \*\*à deux reprises pour arrêter le serveur web en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="ff084-454">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="ff084-455">Sinon, ouvrez une fenêtre Git Bash, ou une invite système Node.JS, et accédez au dossier **Start** du projet.</span><span class="sxs-lookup"><span data-stu-id="ff084-455">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="ff084-456">Bien que le serveur synchronisé au navigateur recharge votre complément dans le volet Office chaque fois que vous apportez une modification à un fichier, y compris le fichier app.js, il ne retranspile pas le code JavaScript. Vous devez donc de nouveau utiliser la commande build afin que les modifications apportées à app.js prennent effet.</span><span class="sxs-lookup"><span data-stu-id="ff084-456">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="ff084-457">Pour ce faire, vous devez arrêter le processus du serveur pour pouvoir obtenir une invite et saisir la commande build.</span><span class="sxs-lookup"><span data-stu-id="ff084-457">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="ff084-458">Une fois la commande build exécutée, redémarrez le serveur.</span><span class="sxs-lookup"><span data-stu-id="ff084-458">After the build, you restart the server.</span></span> <span data-ttu-id="ff084-459">Les prochaines étapes vous permettent d’effectuer ce processus.</span><span class="sxs-lookup"><span data-stu-id="ff084-459">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="ff084-460">Exécutez la commande `npm run build` pour transpiler votre code source ES6 vers une version antérieure de JavaScript prise en charge par Internet Explorer (qui est utilisé en arrière-plan par Excel pour exécuter les compléments Excel).</span><span class="sxs-lookup"><span data-stu-id="ff084-460">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="ff084-461">Exécutez la commande `npm start` pour démarrer un serveur web en cours d’exécution sur localhost.</span><span class="sxs-lookup"><span data-stu-id="ff084-461">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="ff084-462">Recharger le volet Office en le fermant, puis, dans le menu **Accueil**, sélectionnez **Afficher le volet des pages** pour rouvrir le complément.</span><span class="sxs-lookup"><span data-stu-id="ff084-462">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="ff084-463">Sélectionnez le bouton **Boîte de dialogue Ouvrir** dans le volet Office.</span><span class="sxs-lookup"><span data-stu-id="ff084-463">Choose the **Open Dialog** button in the task pane.</span></span>

6. <span data-ttu-id="ff084-464">Lorsque la boîte de dialogue est ouverte, faites-la glisser et redimensionnez-la.</span><span class="sxs-lookup"><span data-stu-id="ff084-464">While the dialog is open, drag it and resize it.</span></span> <span data-ttu-id="ff084-465">Notez que vous pouvez interagir avec la feuille de calcul, appuyez sur les autres boutons dans le volet Office, mais vous ne pouvez pas lancer une deuxième boîte de dialogue à partir de la même page de volet de tâches.</span><span class="sxs-lookup"><span data-stu-id="ff084-465">Note that you can interact with the worksheet and press other buttons on the task pane, but you cannot launch a second dialog from the same task pane page.</span></span>

7. <span data-ttu-id="ff084-466">Dans la boîte de dialogue, entrez un nom et appuyez sur **OK**.</span><span class="sxs-lookup"><span data-stu-id="ff084-466">In the dialog, enter a name and choose **OK**.</span></span> <span data-ttu-id="ff084-467">Ce nom apparaît sur le volet Office et la boîte de dialogue se ferme.</span><span class="sxs-lookup"><span data-stu-id="ff084-467">The name appears on the task pane and the dialog closes.</span></span>

8. <span data-ttu-id="ff084-468">Si vous le souhaitez, vous pouvez commenter la ligne `dialog.close();` dans la fonction `processMessage`.</span><span class="sxs-lookup"><span data-stu-id="ff084-468">Optionally, comment out the line `dialog.close();` in the `processMessage` function.</span></span> <span data-ttu-id="ff084-469">Ensuite, répétez les étapes de cette section.</span><span class="sxs-lookup"><span data-stu-id="ff084-469">Then repeat the steps of this section.</span></span> <span data-ttu-id="ff084-470">La boîte de dialogue reste ouverte et vous pouvez modifier le nom.</span><span class="sxs-lookup"><span data-stu-id="ff084-470">The dialog stays open and you can change the name.</span></span> <span data-ttu-id="ff084-471">Vous pouvez la fermer manuellement en appuyant sur la croix (**X**) en haut à droite.</span><span class="sxs-lookup"><span data-stu-id="ff084-471">You can close it manually by pressing the **X** button in the upper right corner.</span></span>

    ![Didacticiel Excel- Boîte de dialogue](../images/excel-tutorial-dialog-open.png)

## <a name="next-steps"></a><span data-ttu-id="ff084-473">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="ff084-473">Next steps</span></span>

<span data-ttu-id="ff084-474">Ce didacticiel vous apprend à créer un complément Excel qui interagit avec des tableaux, des graphiques (chart), des feuilles de calcul et des boîtes de dialogue dans un classeur Excel.</span><span class="sxs-lookup"><span data-stu-id="ff084-474">In this tutorial, you've created an Excel task pane add-in that interacts with tables, charts, worksheets, and dialogs in an Excel workbook.</span></span> <span data-ttu-id="ff084-475">Pour en savoir plus sur le développement des complément Excel, passez à l’article suivant :</span><span class="sxs-lookup"><span data-stu-id="ff084-475">To learn more about developing Outlook add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="ff084-476">Présentation des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="ff084-476">Excel add-ins overview</span></span>](../excel/excel-add-ins-overview.md)
