---
title: Utilisation de plages à l’aide de l’API JavaScript pour Excel
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 246b882a921b5a43ca747238262af7c4b23c97ee
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459167"
---
# <a name="work-with-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="ccab1-102">Utilisation de plages à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="ccab1-102">Work with ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="ccab1-p101">Cet article fournit des exemples de codes qui montrent comment effectuer des tâches courantes avec des plages à l’aide de l’API JavaScript pour Excel. Pour obtenir la liste complète des propriétés et méthodes que l’objet **Range** prend en charge , voir [l’objet Range (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="ccab1-p101">This article provides code samples that show how to perform common tasks with ranges using the Excel JavaScript API. For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span></span>

## <a name="get-a-range"></a><span data-ttu-id="ccab1-105">Obtenir une plage</span><span class="sxs-lookup"><span data-stu-id="ccab1-105">Get a range</span></span>

<span data-ttu-id="ccab1-106">Les exemples suivants montrent les différentes façons d’obtenir une référence à une plage dans une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="ccab1-106">The following examples show different ways to get a reference to a range within a worksheet.</span></span>

### <a name="get-range-by-address"></a><span data-ttu-id="ccab1-107">Obtenir une plage en fonction d’une adresse</span><span class="sxs-lookup"><span data-stu-id="ccab1-107">Get range by address</span></span>

<span data-ttu-id="ccab1-108">L’exemple de code suivant obtient la plage ayant l’adresse **B2 : B5** à partir de la feuille de calcul nommée **Sample**, charge sa propriété **address** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="ccab1-108">The following code sample gets the range with address **B2:B5** from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

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

### <a name="get-range-by-name"></a><span data-ttu-id="ccab1-109">Obtenir une plage en fonction d’un nom</span><span class="sxs-lookup"><span data-stu-id="ccab1-109">Get range by name</span></span>

<span data-ttu-id="ccab1-110">L’exemple de code suivant obtient la plage nommée **MyRange** à partir de la feuille de calcul nommée **Sample**, charge sa propriété **address** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="ccab1-110">The following code sample gets the range named **MyRange** from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

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

### <a name="get-used-range"></a><span data-ttu-id="ccab1-111">Obtenir une plage utilisée</span><span class="sxs-lookup"><span data-stu-id="ccab1-111">Get used range</span></span>

<span data-ttu-id="ccab1-p102">L’exemple de code suivant obtient la plage utilisée dans la feuille de calcul nommée **Sample**charge sa propriété **address** et écrit un message dans la console. La plage utilisée est la plus petite plage qui englobe des cellules dans la feuille de calcul qui ont une valeur ou une mise en forme attribuée. Si la feuille de calcul entière est vide, la méthode **getUsedRange()** renvoie une plage qui comprend uniquement la cellule en haut à gauche dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="ccab1-p102">The following code sample gets the used range from the worksheet named **Sample**, loads its **address** property, and writes a message to the console. The used range is the smallest range that encompasses any cells in the worksheet that have a value or formatting assigned to them. If the entire worksheet is blank, the **getUsedRange()** method returns a range that consists of only the top-left cell in the worksheet.</span></span>

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

### <a name="get-entire-range"></a><span data-ttu-id="ccab1-115">Obtenir l’intégralité d’une plage</span><span class="sxs-lookup"><span data-stu-id="ccab1-115">Get entire range</span></span>

<span data-ttu-id="ccab1-116">L’exemple de code suivant obtient l’intégralité de la plage de la feuille de calcul à partir de la feuille de calcul nommée **Sample**, charge sa propriété **address** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="ccab1-116">The following code sample gets the entire worksheet range from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

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

## <a name="insert-a-range-of-cells"></a><span data-ttu-id="ccab1-117">Insérer une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="ccab1-117">Insert a range of cells</span></span>

<span data-ttu-id="ccab1-118">L’exemple de code suivant insère une plage de cellules dans l’emplacement **B4:E4** et déplace les autres cellules vers le bas pour laisser de l’espace pour les nouvelles cellules.</span><span class="sxs-lookup"><span data-stu-id="ccab1-118">The following code sample inserts a range of cells in location **B4:E4** and shifts other cells down to provide space for the new cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);
    
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="ccab1-119">**Données avant l’insertion de la plage**</span><span class="sxs-lookup"><span data-stu-id="ccab1-119">**Data before range is inserted**</span></span>

![Données dans Excel avant l’insertion de la plage](../images/excel-ranges-start.png)

<span data-ttu-id="ccab1-121">**Données après l’insertion de la plage**</span><span class="sxs-lookup"><span data-stu-id="ccab1-121">**Data after range is inserted**</span></span>

![Données dans Excel après l’insertion de plage](../images/excel-ranges-after-insert.png)

## <a name="clear-a-range-of-cells"></a><span data-ttu-id="ccab1-123">Effacer une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="ccab1-123">Clear a range of cells</span></span>

<span data-ttu-id="ccab1-124">L’exemple de code suivant efface tout le contenu et la mise en forme des cellules de la plage **E2 : E5**.</span><span class="sxs-lookup"><span data-stu-id="ccab1-124">The following code sample clears all contents and formatting of cells in the range **E2:E5**.</span></span>  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="ccab1-125">**Données avant l’effacement de la plage**</span><span class="sxs-lookup"><span data-stu-id="ccab1-125">**Data before range is cleared**</span></span>

![Données dans Excel avant l’effacement de la plage](../images/excel-ranges-start.png)

<span data-ttu-id="ccab1-127">**Données après l’effacement de plage**</span><span class="sxs-lookup"><span data-stu-id="ccab1-127">**Data after range is cleared**</span></span>

![Données dans Excel après l’effacement de plage](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a><span data-ttu-id="ccab1-129">Supprimer une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="ccab1-129">Delete a range of cells</span></span>

<span data-ttu-id="ccab1-130">L’exemple de code suivant supprime les cellules de la plage **B4:E4** et déplace les autres cellules vers le haut pour remplir l’espace libre suite à la suppression des cellules.</span><span class="sxs-lookup"><span data-stu-id="ccab1-130">The following code sample deletes the cells in the range **B4:E4** and shift other cells up to fill the space that was vacated by the deleted cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="ccab1-131">**Données avant la suppression d’une plage**</span><span class="sxs-lookup"><span data-stu-id="ccab1-131">**Data before range is deleted**</span></span>

![Données dans Excel avant la suppression d’une plage](../images/excel-ranges-start.png)

<span data-ttu-id="ccab1-133">**Données après la suppression d’une plage**</span><span class="sxs-lookup"><span data-stu-id="ccab1-133">**Data after range is deleted**</span></span>

![Données dans Excel après la suppression d’une plage](../images/excel-ranges-after-delete.png)

## <a name="set-the-selected-range"></a><span data-ttu-id="ccab1-135">Définir la plage sélectionnée</span><span class="sxs-lookup"><span data-stu-id="ccab1-135">Set the selected range</span></span>

<span data-ttu-id="ccab1-136">L’exemple de code suivant sélectionne la plage **B2:E6** dans la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="ccab1-136">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="ccab1-137">**Plage sélectionnée  B2:E6**</span><span class="sxs-lookup"><span data-stu-id="ccab1-137">**Selected range B2:E6**</span></span>

![Plage sélectionnée  B2:E6](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="ccab1-139">Obtenir la plage sélectionnée</span><span class="sxs-lookup"><span data-stu-id="ccab1-139">Get the selected range</span></span>

<span data-ttu-id="ccab1-140">L’exemple de code suivant recherche la plage sélectionnée, charge sa propriété **address** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="ccab1-140">The following code sample gets the selected range, loads its **address** property, and writes a message to the console.</span></span> 

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

## <a name="set-values-or-formulas"></a><span data-ttu-id="ccab1-141">Définir des valeurs ou des formules</span><span class="sxs-lookup"><span data-stu-id="ccab1-141">Set values or formulas</span></span>

<span data-ttu-id="ccab1-142">Les exemples suivants indiquent comment définir des valeurs et des formules pour une cellule unique ou une plage de cellules.</span><span class="sxs-lookup"><span data-stu-id="ccab1-142">The following examples show how to set values and formulas for a single cell or a range of cells.</span></span>

### <a name="set-value-for-a-single-cell"></a><span data-ttu-id="ccab1-143">Définir une valeur pour une cellule unique</span><span class="sxs-lookup"><span data-stu-id="ccab1-143">Set value for a single cell</span></span>

<span data-ttu-id="ccab1-144">L’exemple de code suivant définit la valeur de la cellule **C3** à « 5 », puis définit la largeur des colonnes pour mieux s’adapter aux données.</span><span class="sxs-lookup"><span data-stu-id="ccab1-144">The following code sample sets the value of cell **C3** to "5" and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="ccab1-145">**Données avant la mise à jour de la valeur de la cellule**</span><span class="sxs-lookup"><span data-stu-id="ccab1-145">**Data before cell value is updated**</span></span>

![Données dans Excel avant la mise à jour de la valeur de la cellule](../images/excel-ranges-set-start.png)

<span data-ttu-id="ccab1-147">**Données après la mise à jour de la valeur de la cellule**</span><span class="sxs-lookup"><span data-stu-id="ccab1-147">**Data after cell value is updated**</span></span>

![Données dans Excel après la mise à jour de la valeur de la cellule](../images/excel-ranges-set-cell-value.png)

### <a name="set-values-for-a-range-of-cells"></a><span data-ttu-id="ccab1-149">Définir des valeurs pour une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="ccab1-149">Set values for a range of cells</span></span>

<span data-ttu-id="ccab1-150">L’exemple de code suivant définit les valeurs des cellules de la plage **B5:D5**, puis définit la largeur des colonnes pour mieux s’adapter aux données.</span><span class="sxs-lookup"><span data-stu-id="ccab1-150">The following code sample sets values for the cells in the range **B5:D5** and then sets the width of the columns to best fit the data.</span></span>

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

<span data-ttu-id="ccab1-151">**Données avant la mise à jour des valeurs des cellules**</span><span class="sxs-lookup"><span data-stu-id="ccab1-151">**Data before cell values are updated**</span></span>

![Données dans Excel avant la mise à jour des valeurs des cellules](../images/excel-ranges-set-start.png)

<span data-ttu-id="ccab1-153">**Données après la mise à jour des valeurs des cellules**</span><span class="sxs-lookup"><span data-stu-id="ccab1-153">**Data after cell values are updated**</span></span>

![Données dans Excel après la mise à jour des valeurs des cellules](../images/excel-ranges-set-cell-values.png)

### <a name="set-formula-for-a-single-cell"></a><span data-ttu-id="ccab1-155">Définir la formule d’une cellule unique</span><span class="sxs-lookup"><span data-stu-id="ccab1-155">Set formula for a single cell</span></span>

<span data-ttu-id="ccab1-156">L’exemple de code suivant définit une formule pour la cellule **E3**, puis définit la largeur des colonnes pour mieux s’adapter aux données.</span><span class="sxs-lookup"><span data-stu-id="ccab1-156">The following code sample sets a formula for cell **E3** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="ccab1-157">**Données avant la définition de la formule de la cellule**</span><span class="sxs-lookup"><span data-stu-id="ccab1-157">**Data before cell formula is set**</span></span>

![Données dans Excel avant la définition de la formule de la cellule](../images/excel-ranges-start-set-formula.png)

<span data-ttu-id="ccab1-159">**Données après la définition de la formule de la cellule**</span><span class="sxs-lookup"><span data-stu-id="ccab1-159">**Data after cell formula is set**</span></span>

![Données dans Excel après la définition de la formule de la cellule](../images/excel-ranges-set-formula.png)

### <a name="set-formulas-for-a-range-of-cells"></a><span data-ttu-id="ccab1-161">Définir des formules pour une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="ccab1-161">Set formulas for a range of cells</span></span>

<span data-ttu-id="ccab1-162">L’exemple de code ci-dessous définit des formules pour les cellules de la plage **E2:E6**, puis définit la largeur des colonnes pour mieux s’adapter aux données.</span><span class="sxs-lookup"><span data-stu-id="ccab1-162">The following code sample sets formulas for cells in the range **E2:E6** and then sets the width of the columns to best fit the data.</span></span>

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

<span data-ttu-id="ccab1-163">**Données avant la définition des formules des cellules**</span><span class="sxs-lookup"><span data-stu-id="ccab1-163">**Data before cell formulas are set**</span></span>

![Données dans Excel avant la définition des formules des cellules](../images/excel-ranges-start-set-formula.png)

<span data-ttu-id="ccab1-165">**Données après la définition des formules des cellules**</span><span class="sxs-lookup"><span data-stu-id="ccab1-165">**Data after cell formulas are set**</span></span>

![Données dans Excel après la définition des formules des cellules](../images/excel-ranges-set-formulas.png)

## <a name="get-values-text-or-formulas"></a><span data-ttu-id="ccab1-167">Obtenir des valeurs, du texte ou des formules</span><span class="sxs-lookup"><span data-stu-id="ccab1-167">Get values, text, or formulas</span></span>

<span data-ttu-id="ccab1-168">Ces exemples montrent comment obtenir des valeurs, du texte et des formules à partir d’une plage de cellules.</span><span class="sxs-lookup"><span data-stu-id="ccab1-168">These examples show how to get values, text, and formulas from a range of cells.</span></span>

### <a name="get-values-from-a-range-of-cells"></a><span data-ttu-id="ccab1-169">Obtenir des valeurs à partir d’une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="ccab1-169">Get values from a range of cells</span></span>

<span data-ttu-id="ccab1-p103">L’exemple de code suivant obtient la plage **B2:E6**charge sa propriété  **values** et écrit les valeurs dans la console. La propriété **values** d'une plage indique les valeurs brutes que contiennent les cellules. Même si certaines cellules d’une plage contiennent des formules, la propriété **values** de la plage indique les valeurs brutes pour ces cellules, et non les formules.</span><span class="sxs-lookup"><span data-stu-id="ccab1-p103">The following code sample gets the range **B2:E6**, loads its **values** property, and writes the values to the console. The **values** property of a range specifies the raw values that the cells contain. Even if some cells in a range contain formulas, the **values** property of the range specifies the raw values for those cells, not any of the formulas.</span></span>

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

<span data-ttu-id="ccab1-173">**Données de la plage (les valeurs de la colonne E sont le résultat des formules)**</span><span class="sxs-lookup"><span data-stu-id="ccab1-173">**Data in range (values in column E are a result of formulas)**</span></span>

![Données dans Excel après la définition des formules des cellules](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="ccab1-175">**range.values (comme consigné dans la console par l’exemple de code ci-dessus)**</span><span class="sxs-lookup"><span data-stu-id="ccab1-175">**range.values (as logged to the console by the code sample above)**</span></span>

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

### <a name="get-text-from-a-range-of-cells"></a><span data-ttu-id="ccab1-176">Obtenir du texte à partir d’une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="ccab1-176">Get text from a range of cells</span></span>

<span data-ttu-id="ccab1-p104">L’exemple de code suivant obtient la plage **B2:E6**, charge sa propriété **text** et écrit dans la console.  La propriété **text** d’une plage indique les valeurs d'affichage pour les cellules de la plage. Même si certaines cellules d’une plage contiennent des formules, la propriété **text** de la plage indique les valeurs d'affichage pour ces cellules, et non les formules.</span><span class="sxs-lookup"><span data-stu-id="ccab1-p104">The following code sample gets the range **B2:E6**, loads its **text** property, and writes it to the console.  The **text** property of a range specifies the display values for cells in the range. Even if some cells in a range contain formulas, the **text** property of the range specifies the display values for those cells, not any of the formulas.</span></span>

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

<span data-ttu-id="ccab1-180">**Données de la plage (les valeurs de la colonne E sont le résultat des formules)**</span><span class="sxs-lookup"><span data-stu-id="ccab1-180">**Data in range (values in column E are a result of formulas)**</span></span>

![Données dans Excel après la définition des formules des cellules](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="ccab1-182">**range.text (comme consigné dans la console par l’exemple de code ci-dessus)**</span><span class="sxs-lookup"><span data-stu-id="ccab1-182">**range.text (as logged to the console by the code sample above)**</span></span>

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

### <a name="get-formulas-from-a-range-of-cells"></a><span data-ttu-id="ccab1-183">Obtenir des formules à partir d’une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="ccab1-183">Get formulas from a range of cells</span></span>

<span data-ttu-id="ccab1-p105">L’exemple de code suivant obtient la plage **B2:E6**, charge sa propriété **formulas** et écrit dans la console.  La propriété **formulas** d’une plage indique les formules des cellules de la plage qui contiennent des formules et les valeurs brutes pour les cellules de la plage qui ne contiennent pas de formules.</span><span class="sxs-lookup"><span data-stu-id="ccab1-p105">The following code sample gets the range **B2:E6**, loads its **formulas** property, and writes it to the console.  The **formulas** property of a range specifies the formulas for cells in the range that contain formulas and the raw values for cells in the range that do not contain formulas.</span></span>

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

<span data-ttu-id="ccab1-186">**Données de la plage (les valeurs de la colonne E sont le résultat des formules)**</span><span class="sxs-lookup"><span data-stu-id="ccab1-186">**Data in range (values in column E are a result of formulas)**</span></span>

![Données dans Excel après la définition des formules des cellules](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="ccab1-188">**range.formulas (comme consigné dans la console par l’exemple de code ci-dessus)**</span><span class="sxs-lookup"><span data-stu-id="ccab1-188">**range.formulas (as logged to the console by the code sample above)**</span></span>

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

## <a name="set-range-format"></a><span data-ttu-id="ccab1-189">Définir le format de plage</span><span class="sxs-lookup"><span data-stu-id="ccab1-189">Set range format</span></span>

<span data-ttu-id="ccab1-190">Les exemples ci-dessous indiquent comment définir la couleur de police, la couleur de remplissage et le format de nombre pour des cellules dans une plage.</span><span class="sxs-lookup"><span data-stu-id="ccab1-190">The following examples show how to set font color, fill color, and number format for cells in a range.</span></span>

### <a name="set-font-color-and-fill-color"></a><span data-ttu-id="ccab1-191">Définir la couleur de police et la couleur de remplissage</span><span class="sxs-lookup"><span data-stu-id="ccab1-191">Set font color and fill color</span></span>

<span data-ttu-id="ccab1-192">L’exemple de code ci-dessous définit la couleur de police et la couleur de remplissage des cellules de la plage **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="ccab1-192">The following code sample sets the font color and fill color for cells in range **B2:E2**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";;
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="ccab1-193">**Données de la plage avant la définition de la couleur de police et de la couleur de remplissage**</span><span class="sxs-lookup"><span data-stu-id="ccab1-193">**Data in range before font color and fill color are set**</span></span>

![Données dans Excel de la plage avant la définition du format](../images/excel-ranges-format-before.png)

<span data-ttu-id="ccab1-195">**Données de la plage après la définition de la couleur de police et de la couleur de remplissage**</span><span class="sxs-lookup"><span data-stu-id="ccab1-195">**Data in range after font color and fill color are set**</span></span>

![Données dans Excel après la définition du format](../images/excel-ranges-format-font-and-fill.png)

### <a name="set-number-format"></a><span data-ttu-id="ccab1-197">Définir le format de nombre</span><span class="sxs-lookup"><span data-stu-id="ccab1-197">Set number format</span></span>

<span data-ttu-id="ccab1-198">L’exemple de code ci-dessous définit le format de nombre des cellules dans la plage **D3:E5**.</span><span class="sxs-lookup"><span data-stu-id="ccab1-198">The following code sample sets the number format for the cells in range **D3:E5**.</span></span>

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

<span data-ttu-id="ccab1-199">**Données de la plage avant la définition du format de nombre**</span><span class="sxs-lookup"><span data-stu-id="ccab1-199">**Data in range before number format is set**</span></span>

![Données dans Excel de la plage avant la définition du format](../images/excel-ranges-format-font-and-fill.png)

<span data-ttu-id="ccab1-201">**Données de la plage après la définition du format de nombre**</span><span class="sxs-lookup"><span data-stu-id="ccab1-201">**Data in range after number format is set**</span></span>

![Données dans Excel après la définition du format](../images/excel-ranges-format-numbers.png)

## <a name="copy-and-paste"></a><span data-ttu-id="ccab1-203">Copier et coller</span><span class="sxs-lookup"><span data-stu-id="ccab1-203">Copy and paste</span></span>

> [!NOTE]
> <span data-ttu-id="ccab1-p106">La fonction copyFrom est actuellement disponible dans la préversion publique (bêta) uniquement. Pour utiliser cette caractéristique, vous devez utiliser la bibliothèque de la version bêta du RDC Office.js : https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Si vous utilisez TypeScript ou si votre éditeur de code utilise des fichiers de définition de type TypeScript pour IntelliSense, utilisez https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span><span class="sxs-lookup"><span data-stu-id="ccab1-p106">The copyFrom function is currently available only in public preview (beta). To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="ccab1-p107">La fonction copyFrom de la plage reproduit le comportement de copier-coller de l’interface utilisateur d’Excel. L'objet de la plage sur lequel copyFrom est sollicité représente la destination. La source de copie est transmise en tant que plage ou adresse de type chaîne représentant une plage. L’exemple de code suivant copie les données de **A1:E1** vers la plage qui commence à **G1** (finalement collées sur **G1:K1**).</span><span class="sxs-lookup"><span data-stu-id="ccab1-p107">Range’s copyFrom function replicates the copy-and-paste behavior of the Excel UI. The range object that copyFrom is called on is the destination. The source to be copied is passed as a range or a string address representing a range. The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range starting at a single cell destination
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="ccab1-211">Range.copyFrom comporte trois paramètres facultatifs.</span><span class="sxs-lookup"><span data-stu-id="ccab1-211">Range.copyFrom has three optional parameters.</span></span>

```ts
copyFrom(sourceRange: Range | string, copyType?: "All" | "Formulas" | "Values" | "Formats", skipBlanks?: boolean, transpose?: boolean): void;
``` 

<span data-ttu-id="ccab1-p108">`copyType` indique quelles données sont copiées de la source vers la destination.`“Formulas”` transfère les formules dans les cellules source et préserve la position relative de ces plages de formules. Aucune entrée sans formule n'est copiée tel quel. `“Values”` copie les valeurs des données et, dans le cas des formules, le résultat des formules.`“Formats”` copie le format de la plage, y compris la police, la couleur et les autres paramètres de format, mais non les valeurs. `”All”` (l'option par défaut) copie les données et le format, en préservant les formules des cellules le cas échéant.</span><span class="sxs-lookup"><span data-stu-id="ccab1-p108">`copyType` specifies what data gets copied from the source to the destination. `“Formulas”` transfers the formulas in the source cells and preserves the relative positioning of those formulas’ ranges. Any non-formula entries are copied as-is. `“Values”` copies the data values and, in the case of formulas, the result of the formula. `“Formats”` copies the formatting of the range, including font, color, and other format settings, but no values. `”All”` (the default option) copies both data and formatting, preserving cells’ formulas if found.</span></span>

<span data-ttu-id="ccab1-p109">`skipBlanks` indique si les cellules vides sont copiées vers la destination. Lorsque c'est le cas, `copyFrom` ignore les cellules vides de la plage source. Les cellules ignorées ne remplacent pas les données existantes de leurs cellules correspondantes dans la plage de destination. La valeur par défaut est fausse.</span><span class="sxs-lookup"><span data-stu-id="ccab1-p109">`skipBlanks` sets whether blank cells are copied into the destination. When true, `copyFrom` skips blank cells in the source range. Skipped cells will not overwrite the existing data of their corresponding cells in the destination range. The default is false.</span></span>

<span data-ttu-id="ccab1-222">L’exemple de code et les images suivants illustrent ce comportement dans un scénario simple.</span><span class="sxs-lookup"><span data-stu-id="ccab1-222">The following code sample and images demonstrate this behavior in a simple scenario.</span></span> 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="ccab1-223">*Avant que la fonction précédente ait été exécutée.*</span><span class="sxs-lookup"><span data-stu-id="ccab1-223">*Before the preceeding function has been run.*</span></span>

![Les données dans Excel avant que la méthode de copie de plage ait été exécutée.](../images/excel-range-copyfrom-skipblanks-before.png)

<span data-ttu-id="ccab1-225">*Une fois que la fonction précédente a été exécutée.*</span><span class="sxs-lookup"><span data-stu-id="ccab1-225">*After the preceeding function has been run.*</span></span>

![Données dans Excel après l’exécution de la méthode de copie de plage.](../images/excel-range-copyfrom-skipblanks-after.png)

<span data-ttu-id="ccab1-p110">`transpose` détermine si les données sont transposées, ce qui signifie que ses lignes et colonnes sont activées, à l’emplacement source. Une plage transposée pivote le long de la diagonale principale, pour que les lignes **1**, **2**et **3** deviennent les colonnes **A**, **B**et **C**.</span><span class="sxs-lookup"><span data-stu-id="ccab1-p110">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location. A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span> 


## <a name="see-also"></a><span data-ttu-id="ccab1-229">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ccab1-229">See also</span></span>

- [<span data-ttu-id="ccab1-230">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="ccab1-230">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)

