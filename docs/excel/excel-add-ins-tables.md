---
title: Utilisation de tableaux à l’aide de l’API JavaScript pour Excel
description: Exemples de code qui montrent comment effectuer des tâches courantes avec des tableaux à l’aide Excel API JavaScript.
ms.date: 07/02/2021
localization_priority: Normal
ms.openlocfilehash: 931a035b5e46a77afe2f56fecb0cd7ef6d134570
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290788"
---
# <a name="work-with-tables-using-the-excel-javascript-api"></a><span data-ttu-id="3005f-103">Utilisation de tableaux à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="3005f-103">Work with tables using the Excel JavaScript API</span></span>

<span data-ttu-id="3005f-104">Cet article fournit des exemples de code qui expliquent comment effectuer des tâches courantes avec des tableaux à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="3005f-104">This article provides code samples that show how to perform common tasks with tables using the Excel JavaScript API.</span></span> <span data-ttu-id="3005f-105">Pour obtenir la liste complète des propriétés et des méthodes qui sont prise en charge par les objets et les `Table` propriétés, voir Table Object `TableCollection` [(interface API JavaScript](/javascript/api/excel/excel.table) pour Excel) et [TableCollection Object (interface API JavaScript](/javascript/api/excel/excel.tablecollection)pour Excel).</span><span class="sxs-lookup"><span data-stu-id="3005f-105">For the complete list of properties and methods that the `Table` and `TableCollection` objects support, see [Table Object (JavaScript API for Excel)](/javascript/api/excel/excel.table) and [TableCollection Object (JavaScript API for Excel)](/javascript/api/excel/excel.tablecollection).</span></span>

## <a name="create-a-table"></a><span data-ttu-id="3005f-106">Créer un tableau</span><span class="sxs-lookup"><span data-stu-id="3005f-106">Create a table</span></span>

<span data-ttu-id="3005f-107">L’exemple de code suivant crée un tableau dans la feuille de calcul nommée **Sample**.</span><span class="sxs-lookup"><span data-stu-id="3005f-107">The following code sample creates a table in the worksheet named **Sample**.</span></span> <span data-ttu-id="3005f-108">Le tableau comporte des en-têtes et contient quatre colonnes et sept lignes de données.</span><span class="sxs-lookup"><span data-stu-id="3005f-108">The table has headers and contains four columns and seven rows of data.</span></span> <span data-ttu-id="3005f-109">Si l’application Excel où le code [](../reference/requirement-sets/excel-api-requirement-sets.md) est en cours d’exécution prend en charge l’ensemble de conditions **requises ExcelApi 1.2**, la largeur des colonnes et la hauteur des lignes sont définies pour mieux s’adapter aux données actuelles du tableau.</span><span class="sxs-lookup"><span data-stu-id="3005f-109">If the Excel application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

> [!NOTE]
> <span data-ttu-id="3005f-110">Pour spécifier un nom pour une table, vous devez d’abord créer la table, puis définir sa propriété, comme `name` illustré dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="3005f-110">To specify a name for a table, you must first create the table and then set its `name` property, as shown in the following example.</span></span>

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

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="3005f-111">**Nouveau tableau**</span><span class="sxs-lookup"><span data-stu-id="3005f-111">**New table**</span></span>

![Nouveau tableau dans Excel.](../images/excel-tables-create.png)

## <a name="add-rows-to-a-table"></a><span data-ttu-id="3005f-113">Ajouter des lignes dans un tableau</span><span class="sxs-lookup"><span data-stu-id="3005f-113">Add rows to a table</span></span>

<span data-ttu-id="3005f-114">L’exemple de code suivant ajoute sept nouvelles lignes au tableau nommé **ExpensesTable** au sein de la feuille de calcul **Sample**.</span><span class="sxs-lookup"><span data-stu-id="3005f-114">The following code sample adds seven new rows to the table named **ExpensesTable** within the worksheet named **Sample**.</span></span> <span data-ttu-id="3005f-115">Les nouvelles lignes sont ajoutées à la fin du tableau.</span><span class="sxs-lookup"><span data-stu-id="3005f-115">The new rows are added to the end of the table.</span></span> <span data-ttu-id="3005f-116">Si l’application Excel où le code [](../reference/requirement-sets/excel-api-requirement-sets.md) est en cours d’exécution prend en charge l’ensemble de conditions **requises ExcelApi 1.2**, la largeur des colonnes et la hauteur des lignes sont définies pour mieux s’adapter aux données actuelles du tableau.</span><span class="sxs-lookup"><span data-stu-id="3005f-116">If the Excel application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

> [!NOTE]
> <span data-ttu-id="3005f-117">La `index` propriété d’un [objet TableRow](/javascript/api/excel/excel.tablerow) indique le numéro d’index de la ligne dans la collection rows du tableau.</span><span class="sxs-lookup"><span data-stu-id="3005f-117">The `index` property of a [TableRow](/javascript/api/excel/excel.tablerow) object indicates the index number of the row within the rows collection of the table.</span></span> <span data-ttu-id="3005f-118">Un objet ne contient pas de propriété qui peut être utilisée comme clé `TableRow` unique pour identifier la `id` ligne.</span><span class="sxs-lookup"><span data-stu-id="3005f-118">A `TableRow` object does not contain an `id` property that can be used as a unique key to identify the row.</span></span>

> [!WARNING]
> <span data-ttu-id="3005f-119">L’ajout de lignes à un tableau à partir d’un add-in de contenu entraîne une fuite de mémoire.</span><span class="sxs-lookup"><span data-stu-id="3005f-119">Adding rows to a table from a content add-in will result in a memory leak.</span></span> <span data-ttu-id="3005f-120">Consultez [GitHub'#1415 problème](https://github.com/OfficeDev/office-js/issues/1415) pour l’état actuel et des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="3005f-120">See [GitHub Issue #1415](https://github.com/OfficeDev/office-js/issues/1415) for current status and additional information.</span></span> 

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

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="3005f-121">**Tableau avec de nouvelles lignes**</span><span class="sxs-lookup"><span data-stu-id="3005f-121">**Table with new rows**</span></span>

![Tableau avec de nouvelles lignes dans Excel.](../images/excel-tables-add-rows.png)

## <a name="add-a-column-to-a-table"></a><span data-ttu-id="3005f-123">Ajouter une colonne à un tableau</span><span class="sxs-lookup"><span data-stu-id="3005f-123">Add a column to a table</span></span>

<span data-ttu-id="3005f-p106">Ces exemples montrent comment ajouter une colonne à un tableau. Le premier exemple remplit la nouvelle colonne avec des valeurs statiques ; le second exemple remplit la nouvelle colonne avec des formules.</span><span class="sxs-lookup"><span data-stu-id="3005f-p106">These examples show how to add a column to a table. The first example populates the new column with static values; the second example populates the new column with formulas.</span></span>

> [!NOTE]
> <span data-ttu-id="3005f-p107">La propriété **index** d’un objet [TableColumn](/javascript/api/excel/excel.tablecolumn) indique le numéro d’index de la colonne dans la collection de colonnes du tableau. La propriété **id** d’un objet **TableColumn** contient une clé unique qui identifie la colonne.</span><span class="sxs-lookup"><span data-stu-id="3005f-p107">The **index** property of a [TableColumn](/javascript/api/excel/excel.tablecolumn) object indicates the index number of the column within the columns collection of the table. The **id** property of a **TableColumn** object contains a unique key that identifies the column.</span></span>

### <a name="add-a-column-that-contains-static-values"></a><span data-ttu-id="3005f-128">Ajouter une colonne qui contient des valeurs statiques</span><span class="sxs-lookup"><span data-stu-id="3005f-128">Add a column that contains static values</span></span>

<span data-ttu-id="3005f-129">L’exemple de code suivant ajoute une nouvelle colonne à la table nommée **ExpensesTable** au sein de la feuille de calcul **Sample**.</span><span class="sxs-lookup"><span data-stu-id="3005f-129">The following code sample adds a new column to the table named **ExpensesTable** within the worksheet named **Sample**.</span></span> <span data-ttu-id="3005f-130">La nouvelle colonne est ajoutée après les colonnes existantes du tableau et contient un en-tête (« Day of the Week ») ainsi que des données pour remplir les cellules de la colonne.</span><span class="sxs-lookup"><span data-stu-id="3005f-130">The new column is added after all existing columns in the table and contains a header ("Day of the Week") as well as data to populate the cells in the column.</span></span> <span data-ttu-id="3005f-131">Si l’application Excel où le code [](../reference/requirement-sets/excel-api-requirement-sets.md) est en cours d’exécution prend en charge l’ensemble de conditions **requises ExcelApi 1.2**, la largeur des colonnes et la hauteur des lignes sont définies pour mieux s’adapter aux données actuelles du tableau.</span><span class="sxs-lookup"><span data-stu-id="3005f-131">If the Excel application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

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

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="3005f-132">**Tableau avec une nouvelle colonne**</span><span class="sxs-lookup"><span data-stu-id="3005f-132">**Table with new column**</span></span>

![Tableau avec une nouvelle colonne dans Excel.](../images/excel-tables-add-column.png)

### <a name="add-a-column-that-contains-formulas"></a><span data-ttu-id="3005f-134">Ajouter une colonne qui contient des formules</span><span class="sxs-lookup"><span data-stu-id="3005f-134">Add a column that contains formulas</span></span>

<span data-ttu-id="3005f-135">L’exemple de code suivant ajoute une nouvelle colonne à la table nommée **ExpensesTable** au sein de la feuille de calcul **Sample**.</span><span class="sxs-lookup"><span data-stu-id="3005f-135">The following code sample adds a new column to the table named **ExpensesTable** within the worksheet named **Sample**.</span></span> <span data-ttu-id="3005f-136">La nouvelle colonne est ajoutée à la fin du tableau, contient un en-tête («Type of the Day ») et utilise une formule pour remplir chaque cellule de données dans la colonne.</span><span class="sxs-lookup"><span data-stu-id="3005f-136">The new column is added to the end of the table, contains a header ("Type of the Day"), and uses a formula to populate each data cell in the column.</span></span> <span data-ttu-id="3005f-137">Si l’application Excel où le code [](../reference/requirement-sets/excel-api-requirement-sets.md) est en cours d’exécution prend en charge l’ensemble de conditions **requises ExcelApi 1.2**, la largeur des colonnes et la hauteur des lignes sont définies pour mieux s’adapter aux données actuelles du tableau.</span><span class="sxs-lookup"><span data-stu-id="3005f-137">If the Excel application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

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

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="3005f-138">**Tableau avec une nouvelle colonne calculée**</span><span class="sxs-lookup"><span data-stu-id="3005f-138">**Table with new calculated column**</span></span>

![Tableau avec une nouvelle colonne calculée dans Excel.](../images/excel-tables-add-calculated-column.png)

## <a name="resize-a-table"></a><span data-ttu-id="3005f-140">Resize a table</span><span class="sxs-lookup"><span data-stu-id="3005f-140">Resize a table</span></span>

<span data-ttu-id="3005f-141">Votre add-in peut resize un tableau sans ajouter de données au tableau ni modifier les valeurs des cellules.</span><span class="sxs-lookup"><span data-stu-id="3005f-141">Your add-in can resize a table without adding data to the table or changing cell values.</span></span> <span data-ttu-id="3005f-142">Pour re tailler un tableau, utilisez la [méthode Table.resize.](/javascript/api/excel/excel.table#resize_newRange_)</span><span class="sxs-lookup"><span data-stu-id="3005f-142">To resize a table, use the [Table.resize](/javascript/api/excel/excel.table#resize_newRange_) method.</span></span> <span data-ttu-id="3005f-143">L’exemple de code suivant montre comment reizer un tableau.</span><span class="sxs-lookup"><span data-stu-id="3005f-143">The following code sample shows how to resize a table.</span></span> <span data-ttu-id="3005f-144">Cet exemple de code utilise **ExpensesTable** de la [section](#create-a-table) Créer un tableau plus tôt dans cet article et définit la nouvelle plage du tableau sur **A1:D20**.</span><span class="sxs-lookup"><span data-stu-id="3005f-144">This code sample uses the **ExpensesTable** from the [Create a table](#create-a-table) section earlier in this article and sets the new range of the table to **A1:D20**.</span></span>

```js
Excel.run(function (context) {
    // Retrieve the worksheet and a table on that worksheet.
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    // Resize the table.
    expensesTable.resize("A1:D20");

    return context.sync();
}).catch(errorHandlerFunction);
```

> [!IMPORTANT]
> <span data-ttu-id="3005f-145">La nouvelle plage du tableau doit chevaucher la plage d’origine et les en-têtes (ou le haut du tableau) doivent se trouver dans la même ligne.</span><span class="sxs-lookup"><span data-stu-id="3005f-145">The new range of the table must overlap with the original range, and the headers (or the top of the table) must be in the same row.</span></span>

<span data-ttu-id="3005f-146">**Tableau après re resize**</span><span class="sxs-lookup"><span data-stu-id="3005f-146">**Table after resize**</span></span> 

![Tableau avec plusieurs lignes vides dans Excel.](../images/excel-tables-resize.png)

## <a name="update-column-name"></a><span data-ttu-id="3005f-148">Mettre à jour un nom de colonne</span><span class="sxs-lookup"><span data-stu-id="3005f-148">Update column name</span></span>

<span data-ttu-id="3005f-149">L’exemple de code suivant remplace le nom de la première colonne du tableau par **Purchase date**.</span><span class="sxs-lookup"><span data-stu-id="3005f-149">The following code sample updates the name of the first column in the table to **Purchase date**.</span></span> <span data-ttu-id="3005f-150">Si l’application Excel où le code [](../reference/requirement-sets/excel-api-requirement-sets.md) est en cours d’exécution prend en charge l’ensemble de conditions **requises ExcelApi 1.2**, la largeur des colonnes et la hauteur des lignes sont définies pour mieux s’adapter aux données actuelles du tableau.</span><span class="sxs-lookup"><span data-stu-id="3005f-150">If the Excel application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var expensesTable = sheet.tables.getItem("ExpensesTable");
    expensesTable.columns.load("items");

    return context.sync()
        .then(function () {
            expensesTable.columns.items[0].name = "Purchase date";

            if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
                sheet.getUsedRange().format.autofitColumns();
                sheet.getUsedRange().format.autofitRows();
            }

            return context.sync();
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="3005f-151">**Tableau avec un nouveau nom de colonne**</span><span class="sxs-lookup"><span data-stu-id="3005f-151">**Table with new column name**</span></span>

![Tableau avec le nouveau nom de colonne dans Excel.](../images/excel-tables-update-column-name.png)

## <a name="get-data-from-a-table"></a><span data-ttu-id="3005f-153">Obtenir des données à partir d’un tableau</span><span class="sxs-lookup"><span data-stu-id="3005f-153">Get data from a table</span></span>

<span data-ttu-id="3005f-154">L’exemple de code suivant lit les données d’un tableau nommé **ExpensesTable** à partir de la feuille de calcul **Sample**, puis génère ces données en dessous du tableau dans la même feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="3005f-154">The following code sample reads data from a table named **ExpensesTable** in the worksheet named **Sample** and then outputs that data below the table in the same worksheet.</span></span>

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

<span data-ttu-id="3005f-155">**Tableau et sortie des données**</span><span class="sxs-lookup"><span data-stu-id="3005f-155">**Table and data output**</span></span>

![Données de tableau dans Excel.](../images/excel-tables-get-data.png)

## <a name="detect-data-changes"></a><span data-ttu-id="3005f-157">Détecter les modifications de données</span><span class="sxs-lookup"><span data-stu-id="3005f-157">Detect data changes</span></span>

<span data-ttu-id="3005f-158">Votre complément peut avoir besoin de réagir aux utilisateurs modifiant les données dans un tableau.</span><span class="sxs-lookup"><span data-stu-id="3005f-158">Your add-in may need to react to users changing the data in a table.</span></span> <span data-ttu-id="3005f-159">Pour détecter ces modifications, vous pouvez [inscrire un gestionnaire d’événements](excel-add-ins-events.md#register-an-event-handler) à l’événement `onChanged` d’un tableau.</span><span class="sxs-lookup"><span data-stu-id="3005f-159">To detect these changes, you can [register an event handler](excel-add-ins-events.md#register-an-event-handler) for the `onChanged` event of a table.</span></span> <span data-ttu-id="3005f-160">Le gestionnaires d’événements de l’événement `onChanged` reçoit un objet [TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs) lorsque l’événement se déclenche.</span><span class="sxs-lookup"><span data-stu-id="3005f-160">Event handlers for the `onChanged` event receive a [TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs) object when the event fires.</span></span>

<span data-ttu-id="3005f-161">L’objet `TableChangedEventArgs` fournit des informations sur les modifications et la source.</span><span class="sxs-lookup"><span data-stu-id="3005f-161">The `TableChangedEventArgs` object provides information about the changes and the source.</span></span> <span data-ttu-id="3005f-162">Puisque `onChanged` se déclenche lorsque le format ou la valeur des données sont modifiés, il peut être utile que votre complément vérifie si les valeurs ont réellement été modifiées.</span><span class="sxs-lookup"><span data-stu-id="3005f-162">Since `onChanged` fires when either the format or value of the data changes, it can be useful to have your add-in check if the values have actually changed.</span></span> <span data-ttu-id="3005f-163">La propriété de `details` regroupe ces informations en tant qu’un [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail).</span><span class="sxs-lookup"><span data-stu-id="3005f-163">The `details` property encapsulates this information as a [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail).</span></span> <span data-ttu-id="3005f-164">L’exemple de code suivant illustre la procédure d’affichage des valeurs et des types d’une cellule qui a été modifiée, avant et après modification.</span><span class="sxs-lookup"><span data-stu-id="3005f-164">The following code sample shows how to display the before and after values and types of a cell that has been changed.</span></span>

```js
// This function would be used as an event handler for the Table.onChanged event.
function onTableChanged(eventArgs) {
    Excel.run(function (context) {
        var details = eventArgs.details;
        var address = eventArgs.address;

        // Print the before and after types and values to the console.
        console.log(`Change at ${address}: was ${details.valueBefore}(${details.valueTypeBefore}),`
            + ` now is ${details.valueAfter}(${details.valueTypeAfter})`);
        return context.sync();
    });
}
```

## <a name="sort-data-in-a-table"></a><span data-ttu-id="3005f-165">Trier des données dans un tableau</span><span class="sxs-lookup"><span data-stu-id="3005f-165">Sort data in a table</span></span>

<span data-ttu-id="3005f-166">L’exemple de code suivant trie les données d’un tableau dans l’ordre décroissant en fonction des valeurs de la quatrième colonne du tableau.</span><span class="sxs-lookup"><span data-stu-id="3005f-166">The following code sample sorts table data in descending order according to the values in the fourth column of the table.</span></span>

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

<span data-ttu-id="3005f-167">**Données de tableau triées par montant (décroissant)**</span><span class="sxs-lookup"><span data-stu-id="3005f-167">**Table data sorted by Amount (descending)**</span></span>

![Données de table triées dans Excel.](../images/excel-tables-sort.png)

<span data-ttu-id="3005f-169">Lorsque les données sont triées dans une feuille de calcul, une notification d’événement est déclenchée.</span><span class="sxs-lookup"><span data-stu-id="3005f-169">When data is sorted in a worksheet, an event notification fires.</span></span> <span data-ttu-id="3005f-170">Pour en savoir plus sur les événements liés au tri et sur la manière dont votre complément peut inscrire des gestionnaires d’événements pour répondre à ces événements, voir [Gérer les événements de tri](excel-add-ins-worksheets.md#handle-sorting-events).</span><span class="sxs-lookup"><span data-stu-id="3005f-170">To learn more about sort-related events and how your add-in can register event handlers to respond to such events, see [Handle sorting events](excel-add-ins-worksheets.md#handle-sorting-events).</span></span>

## <a name="apply-filters-to-a-table"></a><span data-ttu-id="3005f-171">Appliquer des filtres à un tableau</span><span class="sxs-lookup"><span data-stu-id="3005f-171">Apply filters to a table</span></span>

<span data-ttu-id="3005f-p115">L’exemple de code suivant applique des filtres aux colonnes **Amount** et **Category** d’un tableau. Grâce à l’utilisation des filtres, seules les lignes dans lesquelles **Category** est une des valeurs spécifiées et la valeur de **Amount** est inférieure à la valeur moyenne de toutes les lignes sont affichées.</span><span class="sxs-lookup"><span data-stu-id="3005f-p115">The following code sample applies filters to the **Amount** column and the **Category** column within a table. As a result of the filters, only rows where **Category** is one of the specified values and **Amount** is below the average value for all rows is shown.</span></span>

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

<span data-ttu-id="3005f-174">**Données de tableau avec des filtres appliqués pour les colonnes Category et Amount**</span><span class="sxs-lookup"><span data-stu-id="3005f-174">**Table data with filters applied for Category and Amount**</span></span>

![Données de tableau filtrées dans Excel.](../images/excel-tables-filters-apply.png)

## <a name="clear-table-filters"></a><span data-ttu-id="3005f-176">Effacer les filtres du tableau</span><span class="sxs-lookup"><span data-stu-id="3005f-176">Clear table filters</span></span>

<span data-ttu-id="3005f-177">L’exemple de code suivant efface tous les filtres appliqués actuellement sur le tableau.</span><span class="sxs-lookup"><span data-stu-id="3005f-177">The following code sample clears any filters currently applied on the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.clearFilters();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="3005f-178">**Données de tableau sans filtre appliqué**</span><span class="sxs-lookup"><span data-stu-id="3005f-178">**Table data with no filters applied**</span></span>

![Données de tableau non filtrées dans Excel.](../images/excel-tables-filters-clear.png)

## <a name="get-the-visible-range-from-a-filtered-table"></a><span data-ttu-id="3005f-180">Obtenir la plage visible à partir d’une table filtrée</span><span class="sxs-lookup"><span data-stu-id="3005f-180">Get the visible range from a filtered table</span></span>

<span data-ttu-id="3005f-181">L’exemple de code suivant recherche une plage qui contient des données uniquement pour des cellules qui sont actuellement visibles dans le tableau spécifié, et écrit ensuite les valeurs de la plage dans la console.</span><span class="sxs-lookup"><span data-stu-id="3005f-181">The following code sample gets a range that contains data only for cells that are currently visible within the specified table, and then writes the values of that range to the console.</span></span> <span data-ttu-id="3005f-182">Vous pouvez utiliser la méthode comme indiqué ci-dessous pour obtenir le contenu visible d’un tableau chaque fois que des filtres `getVisibleView()` de colonne ont été appliqués.</span><span class="sxs-lookup"><span data-stu-id="3005f-182">You can use the `getVisibleView()` method as shown below to get the visible contents of a table whenever column filters have been applied.</span></span>

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

## <a name="autofilter"></a><span data-ttu-id="3005f-183">Filtre automatique</span><span class="sxs-lookup"><span data-stu-id="3005f-183">AutoFilter</span></span>

<span data-ttu-id="3005f-184">Un complément peut utiliser l’objet[filtre automatique](/javascript/api/excel/excel.autofilter) du tableau pour filtrer des données.</span><span class="sxs-lookup"><span data-stu-id="3005f-184">An add-in can use the table's [AutoFilter](/javascript/api/excel/excel.autofilter) object to filter data.</span></span> <span data-ttu-id="3005f-185">Un `AutoFilter` objet figure la structure de filtre entière d’une tableau ou d’une plage.</span><span class="sxs-lookup"><span data-stu-id="3005f-185">An `AutoFilter` object is the entire filter structure of a table or range.</span></span> <span data-ttu-id="3005f-186">Toutes les opérations de filtrage décrites précédemment dans cet article sont compatibles avec le filtre automatique.</span><span class="sxs-lookup"><span data-stu-id="3005f-186">All of the filter operations discussed earlier in this article are compatible with the auto-filter.</span></span> <span data-ttu-id="3005f-187">Le point d’accès unique rend plus facile l’accès et la gestion de plusieurs filtres.</span><span class="sxs-lookup"><span data-stu-id="3005f-187">The single access point does make it easier to access and manage multiple filters.</span></span>

<span data-ttu-id="3005f-188">L’exemple de code suivant montre le même [filtrage que celui de l’exemple de code antérieur des données](#apply-filters-to-a-table), mais effectué efficacement et entièrement via le filtre automatique.</span><span class="sxs-lookup"><span data-stu-id="3005f-188">The following code sample shows the same [data filtering as the earlier code sample](#apply-filters-to-a-table), but done entirely through the auto-filter.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.autoFilter.apply(expensesTable.getRange(), 2, {
        filterOn: Excel.FilterOn.values,
        values: ["Restaurant", "Groceries"]
    });
    expensesTable.autoFilter.apply(expensesTable.getRange(), 3, {
        filterOn: Excel.FilterOn.dynamic,
        dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage
    });

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="3005f-189">Un `AutoFilter` peut également être appliqué à une plage au niveau de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="3005f-189">An `AutoFilter` can also be applied to a range at the worksheet level.</span></span> <span data-ttu-id="3005f-190">Pour plus d’informations, consultez [Travailler avec des feuilles de calcul avec l’API JavaScript Excel](excel-add-ins-worksheets.md#filter-data).</span><span class="sxs-lookup"><span data-stu-id="3005f-190">See [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#filter-data) for more information.</span></span>

## <a name="format-a-table"></a><span data-ttu-id="3005f-191">Mettre en forme un tableau</span><span class="sxs-lookup"><span data-stu-id="3005f-191">Format a table</span></span>

<span data-ttu-id="3005f-p119">L’exemple de code suivant applique une mise en forme à un tableau. Il indique différentes couleurs de remplissage pour la ligne d’en-tête, le corps, la deuxième ligne et la première colonne du tableau. Pour plus d’informations sur les propriétés que vous pouvez utiliser pour spécifier un format, reportez-vous à la rubrique [Objet RangeFormat (API JavaScript pour Excel)](/javascript/api/excel/excel.rangeformat).</span><span class="sxs-lookup"><span data-stu-id="3005f-p119">The following code sample applies formatting to a table. It specifies different fill colors for the header row of the table, the body of the table, the second row of the table, and the first column of the table. For information about the properties you can use to specify format, see [RangeFormat Object (JavaScript API for Excel)](/javascript/api/excel/excel.rangeformat).</span></span>

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

<span data-ttu-id="3005f-195">**Tableau après application de la mise en forme**</span><span class="sxs-lookup"><span data-stu-id="3005f-195">**Table after formatting is applied**</span></span>

![Tableau après application de la mise en forme Excel.](../images/excel-tables-formatting-after.png)

## <a name="convert-a-range-to-a-table"></a><span data-ttu-id="3005f-197">Convertir une plage en tableau</span><span class="sxs-lookup"><span data-stu-id="3005f-197">Convert a range to a table</span></span>

<span data-ttu-id="3005f-198">L’exemple de code suivant crée une plage de données, puis la convertit en tableau.</span><span class="sxs-lookup"><span data-stu-id="3005f-198">The following code sample creates a range of data and then converts that range to a table.</span></span>

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

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
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

<span data-ttu-id="3005f-199">**Données de la plage (avant la conversion de la plage en tableau)**</span><span class="sxs-lookup"><span data-stu-id="3005f-199">**Data in the range (before the range is converted to a table)**</span></span>

![Données de plage en Excel.](../images/excel-ranges.png)

<span data-ttu-id="3005f-201">**Données du tableau (après la conversion de la plage en tableau)**</span><span class="sxs-lookup"><span data-stu-id="3005f-201">**Data in the table (after the range is converted to a table)**</span></span>

![Données dans la table dans Excel.](../images/excel-tables-from-range.png)

## <a name="import-json-data-into-a-table"></a><span data-ttu-id="3005f-203">Importer des données JSON dans un tableau</span><span class="sxs-lookup"><span data-stu-id="3005f-203">Import JSON data into a table</span></span>

<span data-ttu-id="3005f-204">L’exemple de code suivant crée un tableau dans la feuille de calcul nommée **Sample** , puis remplit le tableau à l’aide d’un objet JSON qui définit les deux lignes de données.</span><span class="sxs-lookup"><span data-stu-id="3005f-204">The following code sample creates a table in the worksheet named **Sample** and then populates the table by using a JSON object that defines two rows of data.</span></span> <span data-ttu-id="3005f-205">Si l’application Excel où le code [](../reference/requirement-sets/excel-api-requirement-sets.md) est en cours d’exécution prend en charge l’ensemble de conditions **requises ExcelApi 1.2**, la largeur des colonnes et la hauteur des lignes sont définies pour mieux s’adapter aux données actuelles du tableau.</span><span class="sxs-lookup"><span data-stu-id="3005f-205">If the Excel application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

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

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="3005f-206">**Nouveau tableau**</span><span class="sxs-lookup"><span data-stu-id="3005f-206">**New table**</span></span>

![Nouvelle table à partir de données JSON importées dans Excel.](../images/excel-tables-create-from-json.png)

## <a name="see-also"></a><span data-ttu-id="3005f-208">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3005f-208">See also</span></span>

- [<span data-ttu-id="3005f-209">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="3005f-209">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
