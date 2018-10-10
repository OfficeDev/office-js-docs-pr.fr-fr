---
title: Utilisation de tableaux à l’aide de l’API JavaScript pour Excel
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 1e8c71f34de7a295fcac8e5ea6a4fff5cae4fdcf
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459160"
---
# <a name="work-with-tables-using-the-excel-javascript-api"></a><span data-ttu-id="d47ae-102">Utilisation de tableaux à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="d47ae-102">Work with tables using the Excel JavaScript API</span></span>

<span data-ttu-id="d47ae-103">Cet article fournit des exemples de code qui expliquent comment effectuer des tâches courantes avec des tableaux à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="d47ae-103">This article provides code samples that show how to perform common tasks with tables using the Excel JavaScript API.</span></span> <span data-ttu-id="d47ae-104">Pour obtenir une liste complète des propriétés et des méthodes prises en charge par les objets **Table** et **TableCollection**, reportez-vous à la rubrique [Objet Table (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.table?view=office-js) et [Objet TableCollection (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="d47ae-104">For the complete list of properties and methods that the **Table** and **TableCollection** objects support, see [Table Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.table?view=office-js) and [TableCollection Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection?view=office-js).</span></span>

## <a name="create-a-table"></a><span data-ttu-id="d47ae-105">Créer un tableau</span><span class="sxs-lookup"><span data-stu-id="d47ae-105">Create a table</span></span>

<span data-ttu-id="d47ae-106">L’exemple de code suivant crée un tableau dans la feuille de calcul nommée **Sample**.</span><span class="sxs-lookup"><span data-stu-id="d47ae-106">The following code sample creates a table in the worksheet named **Sample**.</span></span> <span data-ttu-id="d47ae-107">Le tableau comporte des en-têtes et contient quatre colonnes et sept lignes de données.</span><span class="sxs-lookup"><span data-stu-id="d47ae-107">The table has headers and contains four columns and seven rows of data.</span></span> <span data-ttu-id="d47ae-108">Si l’application hôte Excel dans laquelle le code est en cours d’exécution prend en charge [l’ensemble de conditions requises](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**, la largeur des colonnes et la hauteur des lignes sont définies pour s’ajuster au mieux aux données actuelles du tableau.</span><span class="sxs-lookup"><span data-stu-id="d47ae-108">If the Excel host application where the code is running supports [requirement set](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

> [!NOTE]
> <span data-ttu-id="d47ae-109">Pour spécifier le nom d’un tableau, vous devez d’abord créer le tableau, puis définir sa propriété **name**, comme indiqué dans l’exemple ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="d47ae-109">To specify a name for a table, you must first create the table and then set its **name** property, as shown in the example below.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";

    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add rows to the end of the table*/, [
        ["1/1/2017", "The Phone Company", "Communications", "$120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "$142"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "$27"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "$33"],
        ["1/11/2017", "Bellows College", "Education", "$350"],
        ["1/15/2017", "Trey Research", "Other", "$135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "$97"]
    ]);

    if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d47ae-110">**Nouveau tableau**</span><span class="sxs-lookup"><span data-stu-id="d47ae-110">**New table**</span></span>

![Nouveau tableau dans Excel](../images/excel-tables-create.png)

## <a name="add-rows-to-a-table"></a><span data-ttu-id="d47ae-112">Ajouter des lignes dans un tableau</span><span class="sxs-lookup"><span data-stu-id="d47ae-112">Add rows to a table</span></span>

<span data-ttu-id="d47ae-113">L’exemple de code suivant ajoute sept nouvelles lignes au tableau nommé **ExpensesTable** au sein de la feuille de calcul **Sample**.</span><span class="sxs-lookup"><span data-stu-id="d47ae-113">The following code sample adds seven new rows to the table named **ExpensesTable** within the worksheet named **Sample**.</span></span> <span data-ttu-id="d47ae-114">Les nouvelles lignes sont ajoutées à la fin du tableau.</span><span class="sxs-lookup"><span data-stu-id="d47ae-114">The new rows are added to the end of the table.</span></span> <span data-ttu-id="d47ae-115">Si l’application hôte Excel dans laquelle le code est en cours d’exécution prend en charge [l’ensemble de conditions requises](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**, la largeur des colonnes et la hauteur des lignes sont définies pour s’ajuster au mieux aux données actuelles du tableau.</span><span class="sxs-lookup"><span data-stu-id="d47ae-115">If the Excel host application where the code is running supports [requirement set](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

> [!NOTE]
> <span data-ttu-id="d47ae-116">La propriété **index** d’un objet [TableRow](https://docs.microsoft.com/javascript/api/excel/excel.tablerow?view=office-js) indique le numéro d’index de la ligne dans la collection de lignes du tableau.</span><span class="sxs-lookup"><span data-stu-id="d47ae-116">The **index** property of a [TableRow](https://docs.microsoft.com/javascript/api/excel/excel.tablerow?view=office-js) object indicates the index number of the row within the rows collection of the table.</span></span> <span data-ttu-id="d47ae-117">Un objet **TableRow** ne contient pas de propriété **id** qui peut être utilisée comme clé unique pour identifier la ligne.</span><span class="sxs-lookup"><span data-stu-id="d47ae-117">A **TableRow** object does not contain an **id** property that can be used as a unique key to identify the row.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");       
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.rows.add(null /*add rows to the end of the table*/, [
        ["1/16/2017", "THE PHONE COMPANY", "Communications", "$120"],
        ["1/20/2017", "NORTHWIND ELECTRIC CARS", "Transportation", "$142"],
        ["1/20/2017", "BEST FOR YOU ORGANICS COMPANY", "Groceries", "$27"],
        ["1/21/2017", "COHO VINEYARD", "Restaurant", "$33"],
        ["1/25/2017", "BELLOWS COLLEGE", "Education", "$350"],
        ["1/28/2017", "TREY RESEARCH", "Other", "$135"],
        ["1/31/2017", "BEST FOR YOU ORGANICS COMPANY", "Groceries", "$97"]
    ]);

    if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d47ae-118">**Tableau avec de nouvelles lignes**</span><span class="sxs-lookup"><span data-stu-id="d47ae-118">**Table with new rows**</span></span>

![Tableau avec de nouvelles lignes dans Excel](../images/excel-tables-add-rows.png)

## <a name="add-a-column-to-a-table"></a><span data-ttu-id="d47ae-120">Ajouter une colonne à un tableau</span><span class="sxs-lookup"><span data-stu-id="d47ae-120">Add a column to a table</span></span>

<span data-ttu-id="d47ae-121">Ces exemples montrent comment ajouter une colonne à un tableau.</span><span class="sxs-lookup"><span data-stu-id="d47ae-121">These examples show how to add a column to a table.</span></span> <span data-ttu-id="d47ae-122">Le premier exemple remplit la nouvelle colonne avec des valeurs statiques ; le second exemple remplit la nouvelle colonne avec des formules.</span><span class="sxs-lookup"><span data-stu-id="d47ae-122">The first example populates the new column with static values; the second example populates the new column with formulas.</span></span>

> [!NOTE]
> <span data-ttu-id="d47ae-123">La propriété **index** d’un objet [TableColumn](https://docs.microsoft.com/javascript/api/excel/excel.tablecolumn?view=office-js) indique le numéro d’index de la colonne dans la collection de colonnes du tableau.</span><span class="sxs-lookup"><span data-stu-id="d47ae-123">The **index** property of a [TableColumn](https://docs.microsoft.com/javascript/api/excel/excel.tablecolumn?view=office-js) object indicates the index number of the column within the columns collection of the table.</span></span> <span data-ttu-id="d47ae-124">La propriété **id** d’un objet **TableColumn** contient une clé unique qui identifie la colonne.</span><span class="sxs-lookup"><span data-stu-id="d47ae-124">The **id** property of a **TableColumn** object contains a unique key that identifies the column.</span></span>

### <a name="add-a-column-that-contains-static-values"></a><span data-ttu-id="d47ae-125">Ajouter une colonne qui contient des valeurs statiques</span><span class="sxs-lookup"><span data-stu-id="d47ae-125">Add a column that contains static values</span></span>

<span data-ttu-id="d47ae-126">L’exemple de code suivant ajoute une nouvelle colonne à la table nommée **ExpensesTable** au sein de la feuille de calcul **Sample**.</span><span class="sxs-lookup"><span data-stu-id="d47ae-126">The following code sample adds a new column to the table named **ExpensesTable** within the worksheet named **Sample**.</span></span> <span data-ttu-id="d47ae-127">La nouvelle colonne est ajoutée après les colonnes existantes du tableau et contient un en-tête (« Day of the Week ») ainsi que des données pour remplir les cellules de la colonne.</span><span class="sxs-lookup"><span data-stu-id="d47ae-127">The new column is added after all existing columns in the table and contains a header ("Day of the Week") as well as data to populate the cells in the column.</span></span> <span data-ttu-id="d47ae-128">Si l’application hôte Excel dans laquelle le code est en cours d’exécution prend en charge [l’ensemble de conditions requises](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**, la largeur des colonnes et la hauteur des lignes sont définies pour s’ajuster au mieux aux données actuelles du tableau.</span><span class="sxs-lookup"><span data-stu-id="d47ae-128">If the Excel host application where the code is running supports [requirement set](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");       
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.columns.add(null /*add columns to the end of the table*/, [
        ["Day of the Week"],
        ["Saturday"],
        ["Friday"],
        ["Monday"],
        ["Thursday"],
        ["Sunday"],
        ["Saturday"],
        ["Monday"]
    ]);

    if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d47ae-129">**Tableau avec une nouvelle colonne**</span><span class="sxs-lookup"><span data-stu-id="d47ae-129">**Table with new column**</span></span>

![Tableau avec une nouvelle colonne dans Excel](../images/excel-tables-add-column.png)

### <a name="add-a-column-that-contains-formulas"></a><span data-ttu-id="d47ae-131">Ajouter une colonne qui contient des formules</span><span class="sxs-lookup"><span data-stu-id="d47ae-131">Add a column that contains formulas</span></span>

<span data-ttu-id="d47ae-132">L’exemple de code suivant ajoute une nouvelle colonne à la table nommée **ExpensesTable** au sein de la feuille de calcul **Sample**.</span><span class="sxs-lookup"><span data-stu-id="d47ae-132">The following code sample adds a new column to the table named **ExpensesTable** within the worksheet named **Sample**.</span></span> <span data-ttu-id="d47ae-133">La nouvelle colonne est ajoutée à la fin du tableau, contient un en-tête (« Type of the Day ») et utilise une formule pour remplir chaque cellule de données dans la colonne.</span><span class="sxs-lookup"><span data-stu-id="d47ae-133">The new column is added to the end of the table, contains a header ("Type of the Day"), and uses a formula to populate each data cell in the column.</span></span> <span data-ttu-id="d47ae-134">Si l’application hôte Excel dans laquelle le code est en cours d’exécution prend en charge [l’ensemble de conditions requises](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**, la largeur des colonnes et la hauteur des lignes sont définies pour s’ajuster au mieux aux données actuelles du tableau.</span><span class="sxs-lookup"><span data-stu-id="d47ae-134">If the Excel host application where the code is running supports [requirement set](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.columns.add(null /*add columns to the end of the table*/, [
        ["Type of the Day"],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")']
    ]);

    if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d47ae-135">**Tableau avec une nouvelle colonne calculée**</span><span class="sxs-lookup"><span data-stu-id="d47ae-135">**Table with new calculated column**</span></span>

![Tableau avec une nouvelle colonne calculée dans Excel](../images/excel-tables-add-calculated-column.png)

## <a name="update-column-name"></a><span data-ttu-id="d47ae-137">Mettre à jour un nom de colonne</span><span class="sxs-lookup"><span data-stu-id="d47ae-137">Update column name</span></span>

<span data-ttu-id="d47ae-p109">L’exemple de code suivant remplace le nom de la première colonne du tableau par **Purchase date**. Si l’application hôte Excel dans laquelle le code est en cours d’exécution prend en charge [l’ensemble de conditions requises](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**, la largeur des colonnes et la hauteur des lignes sont définies pour s’ajuster au mieux aux données actuelles du tableau.</span><span class="sxs-lookup"><span data-stu-id="d47ae-p109">The following code sample updates the name of the first column in the table to **Purchase date**. If the Excel host application where the code is running supports [requirement set](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var expensesTable = sheet.tables.getItem("ExpensesTable");
    expensesTable.columns.load("items");

    return context.sync()
        .then(function () {
            expensesTable.columns.items[0].name = "Purchase date";

            if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
                sheet.getUsedRange().format.autofitColumns();
                sheet.getUsedRange().format.autofitRows();
            }

            return context.sync();
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d47ae-140">**Tableau avec un nouveau nom de colonne**</span><span class="sxs-lookup"><span data-stu-id="d47ae-140">**Table with new column name**</span></span>

![Tableau avec un nouveau nom de colonne dans Excel](../images/excel-tables-update-column-name.png)

## <a name="get-data-from-a-table"></a><span data-ttu-id="d47ae-142">Obtenir des données à partir d’un tableau</span><span class="sxs-lookup"><span data-stu-id="d47ae-142">Get data from a table</span></span>

<span data-ttu-id="d47ae-143">L’exemple de code suivant lit les données d’un tableau nommé **ExpensesTable** à partir de la feuille de calcul **Sample**, puis génère ces données en dessous du tableau dans la même feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="d47ae-143">The following code sample reads data from a table named **ExpensesTable** in the worksheet named **Sample** and then outputs that data below the table in the same worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    // Get data from the header row
    var headerRange = expensesTable.getHeaderRowRange().load("values");

    // Get data from the table
    var bodyRange = expensesTable.getDataBodyRange().load("values");

    // Get data from a single column
    var columnRange = expensesTable.columns.getItem("Merchant").getDataBodyRange().load("values");

    // Get data from a single row
    var rowRange = expensesTable.rows.getItemAt(1).load("values");

    // Sync to populate proxy objects with data from Excel
    return context.sync()
        .then(function () {
            var headerValues = headerRange.values;
            var bodyValues = bodyRange.values;
            var merchantColumnValues = columnRange.values;
            var secondRowValues = rowRange.values;

            // Write data from table back to the sheet
            sheet.getRange("A11:A11").values = [["Results"]];
            sheet.getRange("A13:D13").values = headerValues;
            sheet.getRange("A14:D20").values = bodyValues;
            sheet.getRange("B23:B29").values = merchantColumnValues;
            sheet.getRange("A32:D32").values = secondRowValues;

            // Sync to update the sheet in Excel
            return context.sync();
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d47ae-144">**Tableau et sortie des données**</span><span class="sxs-lookup"><span data-stu-id="d47ae-144">**Table and data output**</span></span>

![Données de tableau dans Excel](../images/excel-tables-get-data.png)

## <a name="sort-data-in-a-table"></a><span data-ttu-id="d47ae-146">Trier des données dans un tableau</span><span class="sxs-lookup"><span data-stu-id="d47ae-146">Sort data in a table</span></span>

<span data-ttu-id="d47ae-147">L’exemple de code suivant trie les données d’un tableau dans l’ordre décroissant en fonction des valeurs de la quatrième colonne du tableau.</span><span class="sxs-lookup"><span data-stu-id="d47ae-147">The following code sample sorts table data in descending order according to the values in the fourth column of the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    // Queue a command to sort data by the fourth column of the table (descending)
    var sortRange = expensesTable.getDataBodyRange();
    sortRange.sort.apply([
        {
            key: 3,
            ascending: false,
        },
    ]);

    // Sync to run the queued command in Excel
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d47ae-148">**Données de tableau triées par montant (décroissant)**</span><span class="sxs-lookup"><span data-stu-id="d47ae-148">**Table data sorted by Amount (descending)**</span></span>

![Données de tableau dans Excel](../images/excel-tables-sort.png)

## <a name="apply-filters-to-a-table"></a><span data-ttu-id="d47ae-150">Appliquer des filtres à un tableau</span><span class="sxs-lookup"><span data-stu-id="d47ae-150">Apply filters to a table</span></span>

<span data-ttu-id="d47ae-151">L’exemple de code suivant applique des filtres aux colonnes **Amount** et **Category** d’un tableau.</span><span class="sxs-lookup"><span data-stu-id="d47ae-151">The following code sample applies filters to the **Amount** column and the **Category** column within a table.</span></span> <span data-ttu-id="d47ae-152">Grâce à l’utilisation des filtres, seules les lignes dans lesquelles **Category** est une des valeurs spécifiées et la valeur de **Amount** est inférieure à la valeur moyenne de toutes les lignes sont affichées.</span><span class="sxs-lookup"><span data-stu-id="d47ae-152">As a result of the filters, only rows where **Category** is one of the specified values and **Amount** is below the average value for all rows is shown.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    // Queue a command to apply a filter on the Category column
    filter = expensesTable.columns.getItem("Category").filter;
    filter.apply({
        filterOn: Excel.FilterOn.values,
        values: ["Restaurant", "Groceries"]
    });

    // Queue a command to apply a filter on the Amount column
    var filter = expensesTable.columns.getItem("Amount").filter;
    filter.apply({
        filterOn: Excel.FilterOn.dynamic,
        dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage
    });

    // Sync to run the queued commands in Excel
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d47ae-153">**Données de tableau avec des filtres appliqués pour les colonnes Category et Amount**</span><span class="sxs-lookup"><span data-stu-id="d47ae-153">**Table data with filters applied for Category and Amount**</span></span>

![Données de tableau filtrées dans Excel](../images/excel-tables-filters-apply.png)

## <a name="clear-table-filters"></a><span data-ttu-id="d47ae-155">Effacer les filtres du tableau</span><span class="sxs-lookup"><span data-stu-id="d47ae-155">Clear table filters</span></span>

<span data-ttu-id="d47ae-156">L’exemple de code suivant efface tous les filtres appliqués actuellement sur le tableau.</span><span class="sxs-lookup"><span data-stu-id="d47ae-156">The following code sample clears any filters currently applied on the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.clearFilters();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d47ae-157">**Données de tableau sans filtre appliqué**</span><span class="sxs-lookup"><span data-stu-id="d47ae-157">**Table data with no filters applied**</span></span>

![Données de tableau non filtrées dans Excel](../images/excel-tables-filters-clear.png)

## <a name="get-the-visible-range-from-a-filtered-table"></a><span data-ttu-id="d47ae-159">Obtenir la plage visible à partir d’une table filtrée</span><span class="sxs-lookup"><span data-stu-id="d47ae-159">Get the visible range from a filtered table</span></span>

<span data-ttu-id="d47ae-160">L’exemple de code suivant recherche une plage qui contient des données uniquement pour des cellules qui sont actuellement visibles dans le tableau spécifié, et écrit ensuite les valeurs de la plage dans la console.</span><span class="sxs-lookup"><span data-stu-id="d47ae-160">The following code sample gets a range that contains data only for cells that are currently visible within the specified table, and then writes the values of that range to the console.</span></span> <span data-ttu-id="d47ae-161">Vous pouvez utiliser la méthode **getVisibleView()** comme indiqué ci-dessous pour rechercher le contenu d’un tableau visible dès que les filtres de colonne ont été appliqués.</span><span class="sxs-lookup"><span data-stu-id="d47ae-161">You can use the **getVisibleView()** method as shown below to get the visible contents of a table whenever column filters have been applied.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    var visibleRange = expensesTable.getDataBodyRange().getVisibleView();
    visibleRange.load("values");

    return context.sync()
        .then(function() {
            console.log(visibleRange.values);
        });
}).catch(errorHandlerFunction);
```

## <a name="format-a-table"></a><span data-ttu-id="d47ae-162">Mettre en forme un tableau</span><span class="sxs-lookup"><span data-stu-id="d47ae-162">Format a table</span></span>

<span data-ttu-id="d47ae-163">L’exemple de code suivant applique une mise en forme à un tableau.</span><span class="sxs-lookup"><span data-stu-id="d47ae-163">The following code sample applies formatting to a table.</span></span> <span data-ttu-id="d47ae-164">Il indique différentes couleurs de remplissage pour la ligne d’en-tête, le corps, la deuxième ligne et la première colonne du tableau.</span><span class="sxs-lookup"><span data-stu-id="d47ae-164">It specifies different fill colors for the header row of the table, the body of the table, the second row of the table, and the first column of the table.</span></span> <span data-ttu-id="d47ae-165">Pour plus d’informations sur les propriétés que vous pouvez utiliser pour spécifier un format, reportez-vous à la rubrique [Objet RangeFormat (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeformat?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="d47ae-165">For information about the properties you can use to specify format, see [RangeFormat Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeformat?view=office-js).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.getHeaderRowRange().format.fill.color = "#C70039";
    expensesTable.getDataBodyRange().format.fill.color = "#DAF7A6";
    expensesTable.rows.getItemAt(1).getRange().format.fill.color = "#FFC300";
    expensesTable.columns.getItemAt(0).getDataBodyRange().format.fill.color = "#FFA07A";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d47ae-166">**Tableau après application de la mise en forme**</span><span class="sxs-lookup"><span data-stu-id="d47ae-166">**Table after formatting is applied**</span></span>

![Tableau après application de la mise en forme dans Excel](../images/excel-tables-formatting-after.png)

## <a name="convert-a-range-to-a-table"></a><span data-ttu-id="d47ae-168">Convertir une plage en tableau</span><span class="sxs-lookup"><span data-stu-id="d47ae-168">Convert a range to a table</span></span>

<span data-ttu-id="d47ae-169">L’exemple de code suivant crée une plage de données, puis la convertit en tableau.</span><span class="sxs-lookup"><span data-stu-id="d47ae-169">The following code sample creates a range of data and then converts that range to a table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Define values for the range
    var values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"],
    ["Frames", 5000, 7000, 6544, 4377],
    ["Saddles", 400, 323, 276, 651],
    ["Brake levers", 12000, 8766, 8456, 9812],
    ["Chains", 1550, 1088, 692, 853],
    ["Mirrors", 225, 600, 923, 544],
    ["Spokes", 6005, 7634, 4589, 8765]];

    // Create the range
    var range = sheet.getRange("A1:E7");
    range.values = values;

    if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    // Convert the range to a table
    var expensesTable = sheet.tables.add('A1:E7', true);
    expensesTable.name = "ExpensesTable";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d47ae-170">**Données de la plage (avant la conversion de la plage en tableau)**</span><span class="sxs-lookup"><span data-stu-id="d47ae-170">**Data in the range (before the range is converted to a table)**</span></span>

![Données de la plage dans Excel](../images/excel-ranges.png)

<span data-ttu-id="d47ae-172">**Données du tableau (après la conversion de la plage en tableau)**</span><span class="sxs-lookup"><span data-stu-id="d47ae-172">**Data in the table (after the range is converted to a table)**</span></span>

![Données du tableau dans Excel](../images/excel-tables-from-range.png)

## <a name="import-json-data-into-a-table"></a><span data-ttu-id="d47ae-174">Importer des données JSON dans un tableau</span><span class="sxs-lookup"><span data-stu-id="d47ae-174">Import JSON data into a table</span></span>

<span data-ttu-id="d47ae-175">L’exemple de code suivant crée un tableau dans la feuille de calcul nommée **Sample** , puis remplit le tableau à l’aide d’un objet JSON qui définit les deux lignes de données.</span><span class="sxs-lookup"><span data-stu-id="d47ae-175">The following code sample creates a table in the worksheet named **Sample** and then populates the table by using a JSON object that defines two rows of data.</span></span> <span data-ttu-id="d47ae-176">Si l’application hôte Excel dans laquelle le code est en cours d’exécution prend en charge [l’ensemble de conditions requises](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**, la largeur des colonnes et la hauteur des lignes sont définies pour s’ajuster au mieux aux données actuelles du tableau.</span><span class="sxs-lookup"><span data-stu-id="d47ae-176">If the Excel host application where the code is running supports [requirement set](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    var transactions = [
      {
        "DATE": "1/1/2017",
        "MERCHANT": "The Phone Company",
        "CATEGORY": "Communications",
        "AMOUNT": "$120"
      },
      {
        "DATE": "1/1/2017",
        "MERCHANT": "Southridge Video",
        "CATEGORY": "Entertainment",
        "AMOUNT": "$40"
      }
    ];

    var newData = transactions.map(item =>
        [item.DATE, item.MERCHANT, item.CATEGORY, item.AMOUNT]);

    expensesTable.rows.add(null, newData);

    if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d47ae-177">**Nouveau tableau**</span><span class="sxs-lookup"><span data-stu-id="d47ae-177">**New table**</span></span>

![Nouveau tableau dans Excel](../images/excel-tables-create-from-json.png)

## <a name="see-also"></a><span data-ttu-id="d47ae-179">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="d47ae-179">See also</span></span>

- [<span data-ttu-id="d47ae-180">Concepts fondamentaux de programmation avec l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="d47ae-180">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)

