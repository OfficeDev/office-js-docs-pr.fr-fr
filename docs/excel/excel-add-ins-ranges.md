---
title: Utilisation de plages à l’aide de l’API JavaScript pour Excel (fondamental)
description: Exemples de code qui montrent comment effectuer des tâches courantes avec des plages à l’aide de l’API JavaScript pour Excel.
ms.date: 07/28/2020
localization_priority: Normal
ms.openlocfilehash: 2cb5640fa9c647c0e6d1f8083986455f2b5a59a3
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408480"
---
# <a name="work-with-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="435a3-103">Utilisation de plages à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="435a3-103">Work with ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="435a3-104">Cet article fournit des exemples de code qui expliquent comment effectuer des tâches courantes avec des plages à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="435a3-104">This article provides code samples that show how to perform common tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="435a3-105">Pour obtenir la liste complète des propriétés et des méthodes `Range` prises en charge par l’objet, reportez-vous à la rubrique [objet Range (interface API JavaScript pour Excel)](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="435a3-105">For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

> [!NOTE]
> <span data-ttu-id="435a3-106">Pour plus d’exemples de code qui montrent comment effectuer des tâches plus avancées avec des plages, consultez l’article [Utiliser les plages à l’aide de l’API JavaScript Excel (avancé)](excel-add-ins-ranges-advanced.md).</span><span class="sxs-lookup"><span data-stu-id="435a3-106">For code samples that show how to perform more advanced tasks with ranges, see [Work with ranges using the Excel JavaScript API (advanced)](excel-add-ins-ranges-advanced.md).</span></span>

## <a name="get-a-range"></a><span data-ttu-id="435a3-107">Obtenir une plage</span><span class="sxs-lookup"><span data-stu-id="435a3-107">Get a range</span></span>

<span data-ttu-id="435a3-108">Les exemples suivants montrent les différentes façons d’obtenir une référence à une plage dans une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="435a3-108">The following examples show different ways to get a reference to a range within a worksheet.</span></span>

### <a name="get-range-by-address"></a><span data-ttu-id="435a3-109">Obtenir une plage en fonction d’une adresse</span><span class="sxs-lookup"><span data-stu-id="435a3-109">Get range by address</span></span>

<span data-ttu-id="435a3-110">L’exemple de code suivant obtient la plage avec l’adresse **B2 : C5** à partir de la feuille de calcul nommée **Sample**, charge sa `address` propriété et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="435a3-110">The following code sample gets the range with address **B2:C5** from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:C5");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range B2:C5 is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-range-by-name"></a><span data-ttu-id="435a3-111">Obtenir une plage en fonction d’un nom</span><span class="sxs-lookup"><span data-stu-id="435a3-111">Get range by name</span></span>

<span data-ttu-id="435a3-112">L’exemple de code suivant obtient la plage nommée `MyRange` à partir de la feuille de calcul nommée **Sample**, charge sa `address` propriété et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="435a3-112">The following code sample gets the range named `MyRange` from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("MyRange");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range "MyRange" is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-used-range"></a><span data-ttu-id="435a3-113">Obtenir une plage utilisée</span><span class="sxs-lookup"><span data-stu-id="435a3-113">Get used range</span></span>

<span data-ttu-id="435a3-114">L’exemple de code suivant obtient la plage utilisée à partir de la feuille de calcul nommée **Sample**, charge sa `address` propriété et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="435a3-114">The following code sample gets the used range from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span> <span data-ttu-id="435a3-115">La plage utilisée est la plus petite plage qui englobe toutes les cellules de la feuille de calcul auxquelles une valeur ou un format est affecté.</span><span class="sxs-lookup"><span data-stu-id="435a3-115">The used range is the smallest range that encompasses any cells in the worksheet that have a value or formatting assigned to them.</span></span> <span data-ttu-id="435a3-116">Si la feuille de calcul entière est vide, la `getUsedRange()` méthode renvoie une plage qui se compose uniquement de la cellule supérieure gauche de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="435a3-116">If the entire worksheet is blank, the `getUsedRange()` method returns a range that consists of only the top-left cell in the worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getUsedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the used range in the worksheet is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-entire-range"></a><span data-ttu-id="435a3-117">Obtenir l’intégralité d’une plage</span><span class="sxs-lookup"><span data-stu-id="435a3-117">Get entire range</span></span>

<span data-ttu-id="435a3-118">L’exemple de code suivant obtient la plage entière de la feuille de calcul à partir de la feuille de calcul nommée **Sample**, charge sa `address` propriété et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="435a3-118">The following code sample gets the entire worksheet range from the worksheet named **Sample**, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the entire worksheet range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="insert-a-range-of-cells"></a><span data-ttu-id="435a3-119">Insérer une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="435a3-119">Insert a range of cells</span></span>

<span data-ttu-id="435a3-120">L’exemple de code suivant insère une plage de cellules dans l’emplacement **B4:E4** et déplace les autres cellules vers le bas pour laisser de l’espace pour les nouvelles cellules.</span><span class="sxs-lookup"><span data-stu-id="435a3-120">The following code sample inserts a range of cells in location **B4:E4** and shifts other cells down to provide space for the new cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-inserted"></a><span data-ttu-id="435a3-121">Données avant l’insertion de la plage</span><span class="sxs-lookup"><span data-stu-id="435a3-121">Data before range is inserted</span></span>

![Données dans Excel avant l’insertion de la plage](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a><span data-ttu-id="435a3-123">Données après l’insertion de la plage</span><span class="sxs-lookup"><span data-stu-id="435a3-123">Data after range is inserted</span></span>

![Données dans Excel après l’insertion de plage](../images/excel-ranges-after-insert.png)

## <a name="clear-a-range-of-cells"></a><span data-ttu-id="435a3-125">Effacer une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="435a3-125">Clear a range of cells</span></span>

<span data-ttu-id="435a3-126">L’exemple de code suivant efface tout le contenu et la mise en forme des cellules de la plage **E2 : E5**.</span><span class="sxs-lookup"><span data-stu-id="435a3-126">The following code sample clears all contents and formatting of cells in the range **E2:E5**.</span></span>  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-cleared"></a><span data-ttu-id="435a3-127">Données avant l’effacement de la plage</span><span class="sxs-lookup"><span data-stu-id="435a3-127">Data before range is cleared</span></span>

![Données dans Excel avant l’effacement de la plage](../images/excel-ranges-start.png)

### <a name="data-after-range-is-cleared"></a><span data-ttu-id="435a3-129">Données après l’effacement de plage</span><span class="sxs-lookup"><span data-stu-id="435a3-129">Data after range is cleared</span></span>

![Données dans Excel après l’effacement de plage](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a><span data-ttu-id="435a3-131">Supprimer une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="435a3-131">Delete a range of cells</span></span>

<span data-ttu-id="435a3-132">L’exemple de code suivant supprime les cellules de la plage **B4:E4** et déplace les autres cellules vers le haut pour remplir l’espace libre suite à la suppression des cellules.</span><span class="sxs-lookup"><span data-stu-id="435a3-132">The following code sample deletes the cells in the range **B4:E4** and shift other cells up to fill the space that was vacated by the deleted cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-deleted"></a><span data-ttu-id="435a3-133">Données avant la suppression d’une plage</span><span class="sxs-lookup"><span data-stu-id="435a3-133">Data before range is deleted</span></span>

![Données dans Excel avant la suppression d’une plage](../images/excel-ranges-start.png)

### <a name="data-after-range-is-deleted"></a><span data-ttu-id="435a3-135">Données après la suppression d’une plage</span><span class="sxs-lookup"><span data-stu-id="435a3-135">Data after range is deleted</span></span>

![Données dans Excel après la suppression d’une plage](../images/excel-ranges-after-delete.png)

## <a name="set-the-selected-range"></a><span data-ttu-id="435a3-137">Définir la plage sélectionnée</span><span class="sxs-lookup"><span data-stu-id="435a3-137">Set the selected range</span></span>

<span data-ttu-id="435a3-138">L’exemple de code suivant sélectionne la plage **B2:E6** dans la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="435a3-138">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a><span data-ttu-id="435a3-139">Plage sélectionnée  B2:E6</span><span class="sxs-lookup"><span data-stu-id="435a3-139">Selected range B2:E6</span></span>

![Plage sélectionnée dans Excel](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="435a3-141">Obtenir la plage sélectionnée</span><span class="sxs-lookup"><span data-stu-id="435a3-141">Get the selected range</span></span>

<span data-ttu-id="435a3-142">L’exemple de code suivant obtient la plage sélectionnée, charge sa `address` propriété et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="435a3-142">The following code sample gets the selected range, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the selected range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="set-values-or-formulas"></a><span data-ttu-id="435a3-143">Définir des valeurs ou des formules</span><span class="sxs-lookup"><span data-stu-id="435a3-143">Set values or formulas</span></span>

<span data-ttu-id="435a3-144">Les exemples suivants indiquent comment définir des valeurs et des formules pour une cellule unique ou une plage de cellules.</span><span class="sxs-lookup"><span data-stu-id="435a3-144">The following examples show how to set values and formulas for a single cell or a range of cells.</span></span>

### <a name="set-value-for-a-single-cell"></a><span data-ttu-id="435a3-145">Définir une valeur pour une cellule unique</span><span class="sxs-lookup"><span data-stu-id="435a3-145">Set value for a single cell</span></span>

<span data-ttu-id="435a3-146">L’exemple de code suivant définit la valeur de la cellule **C3** sur « 5 », puis définit la largeur des colonnes pour mieux s’adapter aux données.</span><span class="sxs-lookup"><span data-stu-id="435a3-146">The following code sample sets the value of cell **C3** to "5" and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-value-is-updated"></a><span data-ttu-id="435a3-147">Données avant la mise à jour de la valeur de la cellule</span><span class="sxs-lookup"><span data-stu-id="435a3-147">Data before cell value is updated</span></span>

![Données dans Excel avant la mise à jour de la valeur de la cellule](../images/excel-ranges-set-start.png)

#### <a name="data-after-cell-value-is-updated"></a><span data-ttu-id="435a3-149">Données après la mise à jour de la valeur de la cellule</span><span class="sxs-lookup"><span data-stu-id="435a3-149">Data after cell value is updated</span></span>

![Données dans Excel après la mise à jour de la valeur de la cellule](../images/excel-ranges-set-cell-value.png)

### <a name="set-values-for-a-range-of-cells"></a><span data-ttu-id="435a3-151">Définir des valeurs pour une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="435a3-151">Set values for a range of cells</span></span>

<span data-ttu-id="435a3-152">L’exemple de code suivant définit les valeurs des cellules de la plage **B5:D5**, puis définit la largeur des colonnes pour mieux s’adapter aux données.</span><span class="sxs-lookup"><span data-stu-id="435a3-152">The following code sample sets values for the cells in the range **B5:D5** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["Potato Chips", 10, 1.80],
    ];

    var range = sheet.getRange("B5:D5");
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-values-are-updated"></a><span data-ttu-id="435a3-153">Données avant la mise à jour des valeurs des cellules</span><span class="sxs-lookup"><span data-stu-id="435a3-153">Data before cell values are updated</span></span>

![Données dans Excel avant la mise à jour des valeurs des cellules](../images/excel-ranges-set-start.png)

#### <a name="data-after-cell-values-are-updated"></a><span data-ttu-id="435a3-155">Données après la mise à jour des valeurs des cellules</span><span class="sxs-lookup"><span data-stu-id="435a3-155">Data after cell values are updated</span></span>

![Données dans Excel après la mise à jour des valeurs des cellules](../images/excel-ranges-set-cell-values.png)

### <a name="set-formula-for-a-single-cell"></a><span data-ttu-id="435a3-157">Définir la formule d’une cellule unique</span><span class="sxs-lookup"><span data-stu-id="435a3-157">Set formula for a single cell</span></span>

<span data-ttu-id="435a3-158">L’exemple de code suivant définit une formule pour la cellule **E3**, puis définit la largeur des colonnes pour mieux s’adapter aux données.</span><span class="sxs-lookup"><span data-stu-id="435a3-158">The following code sample sets a formula for cell **E3** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-formula-is-set"></a><span data-ttu-id="435a3-159">Données avant la définition de la formule de la cellule</span><span class="sxs-lookup"><span data-stu-id="435a3-159">Data before cell formula is set</span></span>

![Données dans Excel avant la définition de la formule de la cellule](../images/excel-ranges-start-set-formula.png)

#### <a name="data-after-cell-formula-is-set"></a><span data-ttu-id="435a3-161">Données après la définition de la formule de la cellule</span><span class="sxs-lookup"><span data-stu-id="435a3-161">Data after cell formula is set</span></span>

![Données dans Excel après la définition de la formule de la cellule](../images/excel-ranges-set-formula.png)

### <a name="set-formulas-for-a-range-of-cells"></a><span data-ttu-id="435a3-163">Définir des formules pour une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="435a3-163">Set formulas for a range of cells</span></span>

<span data-ttu-id="435a3-164">L’exemple de code ci-dessous définit des formules pour les cellules de la plage **E2:E6**, puis définit la largeur des colonnes pour mieux s’adapter aux données.</span><span class="sxs-lookup"><span data-stu-id="435a3-164">The following code sample sets formulas for cells in the range **E2:E6** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"]
    ];

    var range = sheet.getRange("E3:E6");
    range.formulas = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-formulas-are-set"></a><span data-ttu-id="435a3-165">Données avant la définition des formules des cellules</span><span class="sxs-lookup"><span data-stu-id="435a3-165">Data before cell formulas are set</span></span>

![Données dans Excel avant la définition des formules des cellules](../images/excel-ranges-start-set-formula.png)

#### <a name="data-after-cell-formulas-are-set"></a><span data-ttu-id="435a3-167">Données après la définition des formules des cellules</span><span class="sxs-lookup"><span data-stu-id="435a3-167">Data after cell formulas are set</span></span>

![Données dans Excel après la définition des formules des cellules](../images/excel-ranges-set-formulas.png)

## <a name="get-values-text-or-formulas"></a><span data-ttu-id="435a3-169">Obtenir des valeurs, du texte ou des formules</span><span class="sxs-lookup"><span data-stu-id="435a3-169">Get values, text, or formulas</span></span>

<span data-ttu-id="435a3-170">Ces exemples montrent comment obtenir des valeurs, du texte et des formules à partir d’une plage de cellules.</span><span class="sxs-lookup"><span data-stu-id="435a3-170">These examples show how to get values, text, and formulas from a range of cells.</span></span>

### <a name="get-values-from-a-range-of-cells"></a><span data-ttu-id="435a3-171">Obtenir des valeurs à partir d’une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="435a3-171">Get values from a range of cells</span></span>

<span data-ttu-id="435a3-172">L’exemple de code suivant obtient la plage **B2 : E6**, charge sa `values` propriété et écrit les valeurs dans la console.</span><span class="sxs-lookup"><span data-stu-id="435a3-172">The following code sample gets the range **B2:E6**, loads its `values` property, and writes the values to the console.</span></span> <span data-ttu-id="435a3-173">La `values` propriété d’une plage spécifie les valeurs brutes contenues dans les cellules.</span><span class="sxs-lookup"><span data-stu-id="435a3-173">The `values` property of a range specifies the raw values that the cells contain.</span></span> <span data-ttu-id="435a3-174">Même si certaines cellules d’une plage contiennent des formules, la `values` propriété de la plage spécifie les valeurs brutes de ces cellules, pas les formules.</span><span class="sxs-lookup"><span data-stu-id="435a3-174">Even if some cells in a range contain formulas, the `values` property of the range specifies the raw values for those cells, not any of the formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("values");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.values, null, 4));
        });
}).catch(errorHandlerFunction);
```

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a><span data-ttu-id="435a3-175">Données de la plage (les valeurs de la colonne E sont le résultat des formules)</span><span class="sxs-lookup"><span data-stu-id="435a3-175">Data in range (values in column E are a result of formulas)</span></span>

![Données dans Excel après la définition des formules des cellules](../images/excel-ranges-set-formulas.png)

#### <a name="rangevalues-as-logged-to-the-console-by-the-code-sample-above"></a><span data-ttu-id="435a3-177">range.values (comme consigné dans la console par l’exemple de code ci-dessus)</span><span class="sxs-lookup"><span data-stu-id="435a3-177">range.values (as logged to the console by the code sample above)</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        15
    ],
    [
        "Coffee",
        1,
        34.5,
        34.5
    ],
    [
        "Chocolate",
        5,
        9.56,
        47.8
    ],
    [
        "",
        "",
        "",
        97.3
    ]
]
```

### <a name="get-text-from-a-range-of-cells"></a><span data-ttu-id="435a3-178">Obtenir du texte à partir d’une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="435a3-178">Get text from a range of cells</span></span>

<span data-ttu-id="435a3-179">L’exemple de code suivant obtient la plage **B2 : E6**, charge sa `text` propriété et l’écrit dans la console.</span><span class="sxs-lookup"><span data-stu-id="435a3-179">The following code sample gets the range **B2:E6**, loads its `text` property, and writes it to the console.</span></span> <span data-ttu-id="435a3-180">La `text` propriété d’une plage spécifie les valeurs d’affichage pour les cellules de la plage.</span><span class="sxs-lookup"><span data-stu-id="435a3-180">The `text` property of a range specifies the display values for cells in the range.</span></span> <span data-ttu-id="435a3-181">Même si certaines cellules d’une plage contiennent des formules, la `text` propriété de la plage spécifie les valeurs d’affichage de ces cellules, et non des formules.</span><span class="sxs-lookup"><span data-stu-id="435a3-181">Even if some cells in a range contain formulas, the `text` property of the range specifies the display values for those cells, not any of the formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("text");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.text, null, 4));
        });
}).catch(errorHandlerFunction);
```

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a><span data-ttu-id="435a3-182">Données de la plage (les valeurs de la colonne E sont le résultat des formules)</span><span class="sxs-lookup"><span data-stu-id="435a3-182">Data in range (values in column E are a result of formulas)</span></span>

![Données dans Excel après la définition des formules des cellules](../images/excel-ranges-set-formulas.png)

#### <a name="rangetext-as-logged-to-the-console-by-the-code-sample-above"></a><span data-ttu-id="435a3-184">range.text (comme consigné dans la console par l’exemple de code ci-dessus)</span><span class="sxs-lookup"><span data-stu-id="435a3-184">range.text (as logged to the console by the code sample above)</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        "2",
        "7.5",
        "15"
    ],
    [
        "Coffee",
        "1",
        "34.5",
        "34.5"
    ],
    [
        "Chocolate",
        "5",
        "9.56",
        "47.8"
    ],
    [
        "",
        "",
        "",
        "97.3"
    ]
]
```

### <a name="get-formulas-from-a-range-of-cells"></a><span data-ttu-id="435a3-185">Obtenir des formules à partir d’une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="435a3-185">Get formulas from a range of cells</span></span>

<span data-ttu-id="435a3-186">L’exemple de code suivant obtient la plage **B2 : E6**, charge sa `formulas` propriété et l’écrit dans la console.</span><span class="sxs-lookup"><span data-stu-id="435a3-186">The following code sample gets the range **B2:E6**, loads its `formulas` property, and writes it to the console.</span></span> <span data-ttu-id="435a3-187">La `formulas` propriété d’une plage spécifie les formules pour les cellules de la plage qui contiennent des formules et les valeurs brutes pour les cellules de la plage qui ne contiennent pas de formules.</span><span class="sxs-lookup"><span data-stu-id="435a3-187">The `formulas` property of a range specifies the formulas for cells in the range that contain formulas and the raw values for cells in the range that do not contain formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("formulas");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.formulas, null, 4));
        });
}).catch(errorHandlerFunction);
```

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a><span data-ttu-id="435a3-188">Données de la plage (les valeurs de la colonne E sont le résultat des formules)</span><span class="sxs-lookup"><span data-stu-id="435a3-188">Data in range (values in column E are a result of formulas)</span></span>

![Données dans Excel après la définition des formules des cellules](../images/excel-ranges-set-formulas.png)

#### <a name="rangeformulas-as-logged-to-the-console-by-the-code-sample-above"></a><span data-ttu-id="435a3-190">range.formulas (comme consigné dans la console par l’exemple de code ci-dessus)</span><span class="sxs-lookup"><span data-stu-id="435a3-190">range.formulas (as logged to the console by the code sample above)</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        "=C3 * D3"
    ],
    [
        "Coffee",
        1,
        34.5,
        "=C4 * D4"
    ],
    [
        "Chocolate",
        5,
        9.56,
        "=C5 * D5"
    ],
    [
        "",
        "",
        "",
        "=SUM(E3:E5)"
    ]
]
```

## <a name="set-range-format"></a><span data-ttu-id="435a3-191">Définir le format de plage</span><span class="sxs-lookup"><span data-stu-id="435a3-191">Set range format</span></span>

<span data-ttu-id="435a3-192">Les exemples ci-dessous indiquent comment définir la couleur de police, la couleur de remplissage et le format de nombre pour des cellules dans une plage.</span><span class="sxs-lookup"><span data-stu-id="435a3-192">The following examples show how to set font color, fill color, and number format for cells in a range.</span></span>

### <a name="set-font-color-and-fill-color"></a><span data-ttu-id="435a3-193">Définir la couleur de police et la couleur de remplissage</span><span class="sxs-lookup"><span data-stu-id="435a3-193">Set font color and fill color</span></span>

<span data-ttu-id="435a3-194">L’exemple de code ci-dessous définit la couleur de police et la couleur de remplissage des cellules de la plage **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="435a3-194">The following code sample sets the font color and fill color for cells in range **B2:E2**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-in-range-before-font-color-and-fill-color-are-set"></a><span data-ttu-id="435a3-195">Données de la plage avant la définition de la couleur de police et de la couleur de remplissage</span><span class="sxs-lookup"><span data-stu-id="435a3-195">Data in range before font color and fill color are set</span></span>

![Données dans Excel de la plage avant la définition de la couleur de police et de la couleur de remplissage](../images/excel-ranges-format-before.png)

#### <a name="data-in-range-after-font-color-and-fill-color-are-set"></a><span data-ttu-id="435a3-197">Données de la plage après la définition de la couleur de police et de la couleur de remplissage</span><span class="sxs-lookup"><span data-stu-id="435a3-197">Data in range after font color and fill color are set</span></span>

![Données dans Excel après la définition du format](../images/excel-ranges-format-font-and-fill.png)

### <a name="set-number-format"></a><span data-ttu-id="435a3-199">Définir le format de nombre</span><span class="sxs-lookup"><span data-stu-id="435a3-199">Set number format</span></span>

<span data-ttu-id="435a3-200">L’exemple de code ci-dessous définit le format de nombre des cellules dans la plage **D3:E5**.</span><span class="sxs-lookup"><span data-stu-id="435a3-200">The following code sample sets the number format for the cells in range **D3:E5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var formats = [
        ["0.00", "0.00"],
        ["0.00", "0.00"],
        ["0.00", "0.00"]
    ];

    var range = sheet.getRange("D3:E5");
    range.numberFormat = formats;

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-in-range-before-number-format-is-set"></a><span data-ttu-id="435a3-201">Données de la plage avant la définition du format de nombre</span><span class="sxs-lookup"><span data-stu-id="435a3-201">Data in range before number format is set</span></span>

![Données dans Excel avant la définition du format de nombre](../images/excel-ranges-format-font-and-fill.png)

#### <a name="data-in-range-after-number-format-is-set"></a><span data-ttu-id="435a3-203">Données de la plage après la définition du format de nombre</span><span class="sxs-lookup"><span data-stu-id="435a3-203">Data in range after number format is set</span></span>

![Données dans Excel après la définition du format de nombre](../images/excel-ranges-format-numbers.png)

## <a name="read-or-write-to-an-unbounded-range"></a><span data-ttu-id="435a3-205">Lire ou écrire dans une plage non liée</span><span class="sxs-lookup"><span data-stu-id="435a3-205">Read or write to an unbounded range</span></span>

### <a name="read-an-unbounded-range"></a><span data-ttu-id="435a3-206">Lire une plage non liée</span><span class="sxs-lookup"><span data-stu-id="435a3-206">Read an unbounded range</span></span>

<span data-ttu-id="435a3-207">Une adresse de plage illimitée est une adresse de plage qui spécifie des colonnes entières ou des lignes entières.</span><span class="sxs-lookup"><span data-stu-id="435a3-207">An unbounded range address is a range address that specifies either entire columns or entire rows.</span></span> <span data-ttu-id="435a3-208">Par exemple :</span><span class="sxs-lookup"><span data-stu-id="435a3-208">For example:</span></span>

- <span data-ttu-id="435a3-209">Adresses de plage composées de colonnes entières :</span><span class="sxs-lookup"><span data-stu-id="435a3-209">Range addresses comprised of entire columns:</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
- <span data-ttu-id="435a3-210">Adresses de plage composées de lignes entières :</span><span class="sxs-lookup"><span data-stu-id="435a3-210">Range addresses comprised of entire rows:</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>

<span data-ttu-id="435a3-p107">Lorsque l’API effectue une demande de récupération d’une plage non liée (par exemple, `getRange('C:C')`), la réponse contient des valeurs `null` pour les propriétés définies au niveau des cellules, telles que `values`, `text`, `numberFormat` et `formula`. Les autres propriétés de la plage, telles que `address` et `cellCount`, contiennent des valeurs valides pour la plage non liée.</span><span class="sxs-lookup"><span data-stu-id="435a3-p107">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`. Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>

### <a name="write-to-an-unbounded-range"></a><span data-ttu-id="435a3-213">Écrire dans une plage non liée</span><span class="sxs-lookup"><span data-stu-id="435a3-213">Write to an unbounded range</span></span>

<span data-ttu-id="435a3-214">Vous ne pouvez pas définir des propriétés au niveau de la cellule telles que `values` , `numberFormat` , et `formula` sur une plage illimitée, car la demande d’entrée est trop volumineuse.</span><span class="sxs-lookup"><span data-stu-id="435a3-214">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on an unbounded range because the input request is too large.</span></span> <span data-ttu-id="435a3-215">Par exemple, l’extrait de code suivant n’est pas valide, car il tente de spécifier `values` pour une plage non liée.</span><span class="sxs-lookup"><span data-stu-id="435a3-215">For example, the following code snippet is not valid because it attempts to specify `values` for an unbounded range.</span></span> <span data-ttu-id="435a3-216">L’API renvoie une erreur si vous tentez de définir des propriétés au niveau de la cellule pour une plage non liée.</span><span class="sxs-lookup"><span data-stu-id="435a3-216">The API returns an error if you attempt to set cell-level properties for an unbounded range.</span></span>

```js
var range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="read-or-write-to-a-large-range"></a><span data-ttu-id="435a3-217">Lire ou écrire dans une grande plage</span><span class="sxs-lookup"><span data-stu-id="435a3-217">Read or write to a large range</span></span>

<span data-ttu-id="435a3-p109">Si une plage contient un grand nombre de cellules, de valeurs, de formats de nombre et/ou de formules, il n’est peut-être pas possible d’exécuter des opérations d’API sur cette plage. L’API essaie toujours d’exécuter au mieux l’opération demandée sur une plage (par exemple, pour extraire ou écrire des données spécifiées), mais essayer d’effectuer des opérations de lecture ou d’écriture pour une grande plage peut provoquer une erreur d’API en raison de l’utilisation des ressources excessive. Pour éviter ces erreurs, nous vous recommandons d’exécuter des opérations de lecture ou d’écriture distinctes pour des sous-ensembles plus petits d’une grande plage, au lieu d’essayer d’exécuter une seule opération de lecture ou d’écriture sur une grande plage.</span><span class="sxs-lookup"><span data-stu-id="435a3-p109">If a range contains a large number of cells, values, number formats, and/or formulas, it may not be possible to run API operations on that range. The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization. To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>

<span data-ttu-id="435a3-221">Pour plus d’informations sur les limitations système, consultez la section « compléments Excel » de [limites de ressources et optimisation des performances pour les compléments Office](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).</span><span class="sxs-lookup"><span data-stu-id="435a3-221">For details on the system limitations, see the "Excel add-ins" section of [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).</span></span>

### <a name="conditional-formatting-of-ranges"></a><span data-ttu-id="435a3-222">Mise en forme conditionnelle de plages</span><span class="sxs-lookup"><span data-stu-id="435a3-222">Conditional formatting of ranges</span></span>

<span data-ttu-id="435a3-223">Des plages peuvent présenter une mise en forme de cellules individuelles en fonction de certaines conditions.</span><span class="sxs-lookup"><span data-stu-id="435a3-223">Ranges can have formats applied to individual cells based on conditions.</span></span> <span data-ttu-id="435a3-224">Pour plus d’informations à ce sujet, consultez l’article [Appliquer une mise en forme conditionnelle à des plages Excel](excel-add-ins-conditional-formatting.md).</span><span class="sxs-lookup"><span data-stu-id="435a3-224">For more information about this, see [Apply conditional formatting to Excel ranges](excel-add-ins-conditional-formatting.md).</span></span>

## <a name="find-a-cell-using-string-matching"></a><span data-ttu-id="435a3-225">Rechercher une cellule en utilisant la correspondance de chaîne</span><span class="sxs-lookup"><span data-stu-id="435a3-225">Find a cell using string matching</span></span>

<span data-ttu-id="435a3-226">L’objet `Range` dispose d’une méthode`find` pour rechercher une chaîne spécifiée dans la plage.</span><span class="sxs-lookup"><span data-stu-id="435a3-226">The `Range` object has a `find` method to search for a specified string within the range.</span></span> <span data-ttu-id="435a3-227">Elle renvoie la plage de la première cellule avec le texte correspondant.</span><span class="sxs-lookup"><span data-stu-id="435a3-227">It returns the range of the first cell with matching text.</span></span> <span data-ttu-id="435a3-228">L’exemple de code suivant trouve la première cellule contenant une valeur égale à la chaîne **Nourriture** et connecte son adresse à la console.</span><span class="sxs-lookup"><span data-stu-id="435a3-228">The following code sample finds the first cell with a value equal to the string **Food** and logs its address to the console.</span></span> <span data-ttu-id="435a3-229">Notez que `find` génère une erreur `ItemNotFound` si la chaîne spécifiée n’existe pas dans la plage.</span><span class="sxs-lookup"><span data-stu-id="435a3-229">Note that `find` throws an `ItemNotFound` error if the specified string doesn't exist in the range.</span></span> <span data-ttu-id="435a3-230">Si vous pensez que la chaîne spécifiée peut ne pas exister dans la plage, utilisez la méthode[findAllOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) à la place, pour que votre code gère ce scénario plus facilement.</span><span class="sxs-lookup"><span data-stu-id="435a3-230">If you expect that the specified string may not exist in the range, use the [findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) method instead, so your code gracefully handles that scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var table = sheet.tables.getItem("ExpensesTable");
    var searchRange = table.getRange();
    var foundRange = searchRange.find("Food", {
        completeMatch: true, // find will match the whole cell value
        matchCase: false, // find will not match case
        searchDirection: Excel.SearchDirection.forward // find will start searching at the beginning of the range
    });

    foundRange.load("address");
    return context.sync()
        .then(function() {
            console.log(foundRange.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="435a3-231">Lorsque la méthode `find` est appelée sur une plage représentant une cellule simple, la feuille de calcul entière est recherchée.</span><span class="sxs-lookup"><span data-stu-id="435a3-231">When the `find` method is called on a range representing a single cell, the entire worksheet is searched.</span></span> <span data-ttu-id="435a3-232">La recherche commence à cette cellule et continue dans la direction spécifiée par `SearchCriteria.searchDirection`, revenant à la ligne à la fin de la feuille de calcul si nécessaire.</span><span class="sxs-lookup"><span data-stu-id="435a3-232">The search begins at that cell and goes in the direction specified by `SearchCriteria.searchDirection`, wrapping around the ends of the worksheet if needed.</span></span>

## <a name="see-also"></a><span data-ttu-id="435a3-233">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="435a3-233">See also</span></span>

- [<span data-ttu-id="435a3-234">Utiliser les plages à l’aide de l’API JavaScript Excel (avancé)</span><span class="sxs-lookup"><span data-stu-id="435a3-234">Work with ranges using the Excel JavaScript API (advanced)</span></span>](excel-add-ins-ranges-advanced.md)
- [<span data-ttu-id="435a3-235">Modèle objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="435a3-235">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
