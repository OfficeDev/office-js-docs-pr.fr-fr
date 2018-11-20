---
title: Utilisation de plages à l’aide de l’API JavaScript pour Excel
description: ''
ms.date: 10/19/2018
ms.openlocfilehash: 9ac2ce808390dce90572aa27f3f8da2bce9cb572
ms.sourcegitcommit: 8b079005eb042035328e89b29bf2ec775dd08a96
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/24/2018
ms.locfileid: "25772248"
---
# <a name="work-with-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="99d34-102">Utilisation de plages à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="99d34-102">Work with ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="99d34-103">Cet article fournit des exemples de code qui expliquent comment effectuer des tâches courantes avec des plages à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="99d34-103">This article provides code samples that show how to perform common tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="99d34-104">Pour obtenir une liste complète des propriétés et des méthodes prises en charge par l’objet **Range**, reportez-vous à la rubrique [Objet Range (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="99d34-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span></span>

## <a name="get-a-range"></a><span data-ttu-id="99d34-105">Obtenir une plage</span><span class="sxs-lookup"><span data-stu-id="99d34-105">Get a range</span></span>

<span data-ttu-id="99d34-106">Les exemples suivants montrent les différentes façons d’obtenir une référence à une plage dans une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="99d34-106">The following examples show different ways to get a reference to a range within a worksheet.</span></span>

### <a name="get-range-by-address"></a><span data-ttu-id="99d34-107">Obtenir une plage en fonction d’une adresse</span><span class="sxs-lookup"><span data-stu-id="99d34-107">Get range by address</span></span>

<span data-ttu-id="99d34-108">L’exemple de code suivant obtient la plage ayant l’adresse **B2 : B5** à partir de la feuille de calcul nommée **Sample**, charge sa propriété **address** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="99d34-108">The following code sample gets the range with address **B2:B5** from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

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

### <a name="get-range-by-name"></a><span data-ttu-id="99d34-109">Obtenir une plage en fonction d’un nom</span><span class="sxs-lookup"><span data-stu-id="99d34-109">Get range by name</span></span>

<span data-ttu-id="99d34-110">L’exemple de code suivant obtient la plage nommée **MyRange** à partir de la feuille de calcul nommée **Sample**, charge sa propriété **address** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="99d34-110">The following code sample gets the range named **MyRange** from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

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

### <a name="get-used-range"></a><span data-ttu-id="99d34-111">Obtenir une plage utilisée</span><span class="sxs-lookup"><span data-stu-id="99d34-111">Get used range</span></span>

<span data-ttu-id="99d34-112">L’exemple de code suivant obtient la plage  utilisée dans la feuille de calcul nommée **Sample**, charge sa propriété **address** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="99d34-112">The following code sample gets the used range from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span> <span data-ttu-id="99d34-113">La plage utilisée est la plus petite plage qui englobe toutes les cellules de la feuille de calcul auxquelles une valeur ou un format est affecté.</span><span class="sxs-lookup"><span data-stu-id="99d34-113">The used range is the smallest range that encompasses any cells in the worksheet that have a value or formatting assigned to them.</span></span> <span data-ttu-id="99d34-114">Si la feuille de calcul entière est vide, la méthode **getUsedRange()** renvoie une plage qui se compose d’uniquement de la cellule en haut à gauche de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="99d34-114">If the entire worksheet is blank, the **getUsedRange()** method returns a range that consists of only the top-left cell in the worksheet.</span></span>

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

### <a name="get-entire-range"></a><span data-ttu-id="99d34-115">Obtenir l’intégralité d’une plage</span><span class="sxs-lookup"><span data-stu-id="99d34-115">Get entire range</span></span>

<span data-ttu-id="99d34-116">L’exemple de code suivant obtient l’intégralité de la plage de la feuille de calcul à partir de la feuille de calcul nommée **Sample**, charge sa propriété **address** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="99d34-116">The following code sample gets the entire worksheet range from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

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

## <a name="insert-a-range-of-cells"></a><span data-ttu-id="99d34-117">Insérer une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="99d34-117">Insert a range of cells</span></span>

<span data-ttu-id="99d34-118">L’exemple de code suivant insère une plage de cellules dans l’emplacement **B4:E4** et déplace les autres cellules vers le bas pour laisser de l’espace pour les nouvelles cellules.</span><span class="sxs-lookup"><span data-stu-id="99d34-118">The following code sample inserts a range of cells in location **B4:E4** and shifts other cells down to provide space for the new cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);
    
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="99d34-119">**Données avant l’insertion de la plage**</span><span class="sxs-lookup"><span data-stu-id="99d34-119">**Data before range is inserted**</span></span>

![Données dans Excel avant l’insertion de la plage](../images/excel-ranges-start.png)

<span data-ttu-id="99d34-121">**Données après l’insertion de la plage**</span><span class="sxs-lookup"><span data-stu-id="99d34-121">**Data after range is inserted**</span></span>

![Données dans Excel après l’insertion de plage](../images/excel-ranges-after-insert.png)

## <a name="clear-a-range-of-cells"></a><span data-ttu-id="99d34-123">Effacer une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="99d34-123">Clear a range of cells</span></span>

<span data-ttu-id="99d34-124">L’exemple de code suivant efface tout le contenu et la mise en forme des cellules de la plage **E2 : E5**.</span><span class="sxs-lookup"><span data-stu-id="99d34-124">The following code sample clears all contents and formatting of cells in the range **E2:E5**.</span></span>  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="99d34-125">**Données avant l’effacement de la plage**</span><span class="sxs-lookup"><span data-stu-id="99d34-125">**Data before range is cleared**</span></span>

![Données dans Excel avant l’effacement de la plage](../images/excel-ranges-start.png)

<span data-ttu-id="99d34-127">**Données après l’effacement de plage**</span><span class="sxs-lookup"><span data-stu-id="99d34-127">**Data after range is cleared**</span></span>

![Données dans Excel après l’effacement de plage](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a><span data-ttu-id="99d34-129">Supprimer une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="99d34-129">Delete a range of cells</span></span>

<span data-ttu-id="99d34-130">L’exemple de code suivant supprime les cellules de la plage **B4:E4** et déplace les autres cellules vers le haut pour remplir l’espace libre suite à la suppression des cellules.</span><span class="sxs-lookup"><span data-stu-id="99d34-130">The following code sample deletes the cells in the range **B4:E4** and shift other cells up to fill the space that was vacated by the deleted cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="99d34-131">**Données avant la suppression d’une plage**</span><span class="sxs-lookup"><span data-stu-id="99d34-131">**Data before range is deleted**</span></span>

![Données dans Excel avant la suppression d’une plage](../images/excel-ranges-start.png)

<span data-ttu-id="99d34-133">**Données après la suppression d’une plage**</span><span class="sxs-lookup"><span data-stu-id="99d34-133">**Data after range is deleted**</span></span>

![Données dans Excel après la suppression d’une plage](../images/excel-ranges-after-delete.png)

## <a name="set-the-selected-range"></a><span data-ttu-id="99d34-135">Définir la plage sélectionnée</span><span class="sxs-lookup"><span data-stu-id="99d34-135">Set the selected range</span></span>

<span data-ttu-id="99d34-136">L’exemple de code suivant sélectionne la plage **B2:E6** dans la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="99d34-136">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="99d34-137">**Plage sélectionnée  B2:E6**</span><span class="sxs-lookup"><span data-stu-id="99d34-137">**Selected range B2:E6**</span></span>

![Plage sélectionnée dans Excel](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="99d34-139">Obtenir la plage sélectionnée</span><span class="sxs-lookup"><span data-stu-id="99d34-139">Get the selected range</span></span>

<span data-ttu-id="99d34-140">L’exemple de code suivant recherche la plage  sélectionnée, charge sa propriété **address** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="99d34-140">The following code sample gets the selected range, loads its **address** property, and writes a message to the console.</span></span> 

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

## <a name="set-values-or-formulas"></a><span data-ttu-id="99d34-141">Définir des valeurs ou des formules</span><span class="sxs-lookup"><span data-stu-id="99d34-141">Set values or formulas</span></span>

<span data-ttu-id="99d34-142">Les exemples suivants indiquent comment définir des valeurs et des formules pour une cellule unique ou une plage de cellules.</span><span class="sxs-lookup"><span data-stu-id="99d34-142">The following examples show how to set values and formulas for a single cell or a range of cells.</span></span>

### <a name="set-value-for-a-single-cell"></a><span data-ttu-id="99d34-143">Définir une valeur pour une cellule unique</span><span class="sxs-lookup"><span data-stu-id="99d34-143">Set value for a single cell</span></span>

<span data-ttu-id="99d34-144">L’exemple de code suivant définit la valeur de la cellule **C3** sur « 5 », puis définit la largeur des colonnes pour mieux s’adapter aux données.</span><span class="sxs-lookup"><span data-stu-id="99d34-144">The following code sample sets the value of cell **C3** to "5" and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="99d34-145">**Données avant la mise à jour de la valeur de la cellule**</span><span class="sxs-lookup"><span data-stu-id="99d34-145">**Data before cell value is updated**</span></span>

![Données dans Excel avant la mise à jour de la valeur de la cellule](../images/excel-ranges-set-start.png)

<span data-ttu-id="99d34-147">**Données après la mise à jour de la valeur de la cellule**</span><span class="sxs-lookup"><span data-stu-id="99d34-147">**Data after cell value is updated**</span></span>

![Données dans Excel après la mise à jour de la valeur de la cellule](../images/excel-ranges-set-cell-value.png)

### <a name="set-values-for-a-range-of-cells"></a><span data-ttu-id="99d34-149">Définir des valeurs pour une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="99d34-149">Set values for a range of cells</span></span>

<span data-ttu-id="99d34-150">L’exemple de code suivant définit les valeurs des cellules de la plage **B5:D5**, puis définit la largeur des colonnes pour mieux s’adapter aux données.</span><span class="sxs-lookup"><span data-stu-id="99d34-150">The following code sample sets values for the cells in the range **B5:D5** and then sets the width of the columns to best fit the data.</span></span>

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

<span data-ttu-id="99d34-151">**Données avant la mise à jour des valeurs des cellules**</span><span class="sxs-lookup"><span data-stu-id="99d34-151">**Data before cell values are updated**</span></span>

![Données dans Excel avant la mise à jour des valeurs des cellules](../images/excel-ranges-set-start.png)

<span data-ttu-id="99d34-153">**Données après la mise à jour des valeurs des cellules**</span><span class="sxs-lookup"><span data-stu-id="99d34-153">**Data after cell values are updated**</span></span>

![Données dans Excel après la mise à jour des valeurs des cellules](../images/excel-ranges-set-cell-values.png)

### <a name="set-formula-for-a-single-cell"></a><span data-ttu-id="99d34-155">Définir la formule d’une cellule unique</span><span class="sxs-lookup"><span data-stu-id="99d34-155">Set formula for a single cell</span></span>

<span data-ttu-id="99d34-156">L’exemple de code suivant définit une formule pour la cellule **E3**, puis définit la largeur des colonnes pour mieux s’adapter aux données.</span><span class="sxs-lookup"><span data-stu-id="99d34-156">The following code sample sets a formula for cell **E3** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="99d34-157">**Données avant la définition de la formule de la cellule**</span><span class="sxs-lookup"><span data-stu-id="99d34-157">**Data before cell formula is set**</span></span>

![Données dans Excel avant la définition de la formule de la cellule](../images/excel-ranges-start-set-formula.png)

<span data-ttu-id="99d34-159">**Données après la définition de la formule de la cellule**</span><span class="sxs-lookup"><span data-stu-id="99d34-159">**Data after cell formula is set**</span></span>

![Données dans Excel après la définition de la formule de la cellule](../images/excel-ranges-set-formula.png)

### <a name="set-formulas-for-a-range-of-cells"></a><span data-ttu-id="99d34-161">Définir des formules pour une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="99d34-161">Set formulas for a range of cells</span></span>

<span data-ttu-id="99d34-162">L’exemple de code ci-dessous définit des formules pour les cellules de la plage **E2:E6**, puis définit la largeur des colonnes pour mieux s’adapter aux données.</span><span class="sxs-lookup"><span data-stu-id="99d34-162">The following code sample sets formulas for cells in the range **E2:E6** and then sets the width of the columns to best fit the data.</span></span>

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

<span data-ttu-id="99d34-163">**Données avant la définition des formules des cellules**</span><span class="sxs-lookup"><span data-stu-id="99d34-163">**Data before cell formulas are set**</span></span>

![Données dans Excel avant la définition des formules des cellules](../images/excel-ranges-start-set-formula.png)

<span data-ttu-id="99d34-165">**Données après la définition des formules des cellules**</span><span class="sxs-lookup"><span data-stu-id="99d34-165">**Data after cell formulas are set**</span></span>

![Données dans Excel après la définition des formules des cellules](../images/excel-ranges-set-formulas.png)

## <a name="get-values-text-or-formulas"></a><span data-ttu-id="99d34-167">Obtenir des valeurs, du texte ou des formules</span><span class="sxs-lookup"><span data-stu-id="99d34-167">Get values, text, or formulas</span></span>

<span data-ttu-id="99d34-168">Ces exemples montrent comment obtenir des valeurs, du texte et des formules à partir d’une plage de cellules.</span><span class="sxs-lookup"><span data-stu-id="99d34-168">These examples show how to get values, text, and formulas from a range of cells.</span></span>

### <a name="get-values-from-a-range-of-cells"></a><span data-ttu-id="99d34-169">Obtenir des valeurs à partir d’une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="99d34-169">Get values from a range of cells</span></span>

<span data-ttu-id="99d34-170">L’exemple de code suivant obtient la plage **B2:E6**, charge la propriété **values** et écrit les valeurs dans la console.</span><span class="sxs-lookup"><span data-stu-id="99d34-170">The following code sample gets the range **B2:E6**, loads its **values** property, and writes the values to the console.</span></span> <span data-ttu-id="99d34-171">La propriété **values** d’une plage spécifie les valeurs brutes contenues dans les cellules.</span><span class="sxs-lookup"><span data-stu-id="99d34-171">The **values** property of a range specifies the raw values that the cells contain.</span></span> <span data-ttu-id="99d34-172">Même si certaines cellules d’une plage contiennent des formules, la propriété **values** de la plage spécifie les valeurs brutes des cellules, et non des formules.</span><span class="sxs-lookup"><span data-stu-id="99d34-172">Even if some cells in a range contain formulas, the **values** property of the range specifies the raw values for those cells, not any of the formulas.</span></span>

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

<span data-ttu-id="99d34-173">**Données de la plage (les valeurs de la colonne E sont le résultat des formules)**</span><span class="sxs-lookup"><span data-stu-id="99d34-173">**Data in range (values in column E are a result of formulas)**</span></span>

![Données dans Excel après la définition des formules des cellules](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="99d34-175">**range.values (comme consigné dans la console par l’exemple de code ci-dessus)**</span><span class="sxs-lookup"><span data-stu-id="99d34-175">**range.values (as logged to the console by the code sample above)**</span></span>

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

### <a name="get-text-from-a-range-of-cells"></a><span data-ttu-id="99d34-176">Obtenir du texte à partir d’une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="99d34-176">Get text from a range of cells</span></span>

<span data-ttu-id="99d34-177">L’exemple de code suivant obtient la plage **B2:E6**, charge sa propriété **text** et l’écrit dans la console.</span><span class="sxs-lookup"><span data-stu-id="99d34-177">The following code sample gets the range **B2:E6**, loads its **text** property, and writes it to the console.</span></span>  <span data-ttu-id="99d34-178">La propriété **text** d’une plage spécifie les valeurs d’affichage pour les cellules de la plage.</span><span class="sxs-lookup"><span data-stu-id="99d34-178">The **text** property of a range specifies the display values for cells in the range.</span></span> <span data-ttu-id="99d34-179">Même si certaines cellules d’une plage contiennent des formules, la propriété **text** de la plage indique les valeurs d’affichage pour ces cellules, et non des formules.</span><span class="sxs-lookup"><span data-stu-id="99d34-179">Even if some cells in a range contain formulas, the **text** property of the range specifies the display values for those cells, not any of the formulas.</span></span>

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

<span data-ttu-id="99d34-180">**Données de la plage (les valeurs de la colonne E sont le résultat des formules)**</span><span class="sxs-lookup"><span data-stu-id="99d34-180">**Data in range (values in column E are a result of formulas)**</span></span>

![Données dans Excel après la définition des formules des cellules](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="99d34-182">**range.text (comme consigné dans la console par l’exemple de code ci-dessus)**</span><span class="sxs-lookup"><span data-stu-id="99d34-182">**range.text (as logged to the console by the code sample above)**</span></span>

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

### <a name="get-formulas-from-a-range-of-cells"></a><span data-ttu-id="99d34-183">Obtenir des formules à partir d’une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="99d34-183">Get formulas from a range of cells</span></span>

<span data-ttu-id="99d34-184">L’exemple de code suivant obtient la plage **B2:E6**, charge sa propriété **formulas** et l’écrit dans la console.</span><span class="sxs-lookup"><span data-stu-id="99d34-184">The following code sample gets the range **B2:E6**, loads its **formulas** property, and writes it to the console.</span></span>  <span data-ttu-id="99d34-185">La propriété **formulas** d’une plage spécifie les formules pour les cellules de la plage contenant des formules et des valeurs brutes pour les cellules de la plage ne contenant pas de formule.</span><span class="sxs-lookup"><span data-stu-id="99d34-185">The **formulas** property of a range specifies the formulas for cells in the range that contain formulas and the raw values for cells in the range that do not contain formulas.</span></span>

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

<span data-ttu-id="99d34-186">**Données de la plage (les valeurs de la colonne E sont le résultat des formules)**</span><span class="sxs-lookup"><span data-stu-id="99d34-186">**Data in range (values in column E are a result of formulas)**</span></span>

![Données dans Excel après la définition des formules des cellules](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="99d34-188">**range.formulas (comme consigné dans la console par l’exemple de code ci-dessus)**</span><span class="sxs-lookup"><span data-stu-id="99d34-188">**range.formulas (as logged to the console by the code sample above)**</span></span>

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

## <a name="set-range-format"></a><span data-ttu-id="99d34-189">Définir le format de plage</span><span class="sxs-lookup"><span data-stu-id="99d34-189">Set range format</span></span>

<span data-ttu-id="99d34-190">Les exemples ci-dessous indiquent comment définir la couleur de police, la couleur de remplissage et le format de nombre pour des cellules dans une plage.</span><span class="sxs-lookup"><span data-stu-id="99d34-190">The following examples show how to set font color, fill color, and number format for cells in a range.</span></span>

### <a name="set-font-color-and-fill-color"></a><span data-ttu-id="99d34-191">Définir la couleur de police et la couleur de remplissage</span><span class="sxs-lookup"><span data-stu-id="99d34-191">Set font color and fill color</span></span>

<span data-ttu-id="99d34-192">L’exemple de code ci-dessous définit la couleur de police et la couleur de remplissage des cellules de la plage **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="99d34-192">The following code sample sets the font color and fill color for cells in range **B2:E2**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";;
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="99d34-193">**Données de la plage avant la définition de la couleur de police et de la couleur de remplissage**</span><span class="sxs-lookup"><span data-stu-id="99d34-193">**Data in range before font color and fill color are set**</span></span>

![Données dans Excel de la plage avant la définition de la couleur de police et de la couleur de remplissage](../images/excel-ranges-format-before.png)

<span data-ttu-id="99d34-195">**Données de la plage après la définition de la couleur de police et de la couleur de remplissage**</span><span class="sxs-lookup"><span data-stu-id="99d34-195">**Data in range after font color and fill color are set**</span></span>

![Données dans Excel après la définition du format](../images/excel-ranges-format-font-and-fill.png)

### <a name="set-number-format"></a><span data-ttu-id="99d34-197">Définir le format de nombre</span><span class="sxs-lookup"><span data-stu-id="99d34-197">Set number format</span></span>

<span data-ttu-id="99d34-198">L’exemple de code ci-dessous définit le format de nombre des cellules dans la plage **D3:E5**.</span><span class="sxs-lookup"><span data-stu-id="99d34-198">The following code sample sets the number format for the cells in range **D3:E5**.</span></span>

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

<span data-ttu-id="99d34-199">**Données de la plage avant la définition du format de nombre**</span><span class="sxs-lookup"><span data-stu-id="99d34-199">**Data in range before number format is set**</span></span>

![Données dans Excel de la plage avant la définition de la couleur de police et de la couleur de remplissage](../images/excel-ranges-format-font-and-fill.png)

<span data-ttu-id="99d34-201">**Données de la plage après la définition du format de nombre**</span><span class="sxs-lookup"><span data-stu-id="99d34-201">**Data in range after number format is set**</span></span>

![Données dans Excel après la définition du format](../images/excel-ranges-format-numbers.png)

### <a name="conditional-formatting-of-ranges"></a><span data-ttu-id="99d34-203">Mise en forme conditionnelle de plages</span><span class="sxs-lookup"><span data-stu-id="99d34-203">Conditional formatting of ranges</span></span>

<span data-ttu-id="99d34-204">Des plages peuvent présenter une mise en forme de cellules individuelles en fonction de certaines conditions.</span><span class="sxs-lookup"><span data-stu-id="99d34-204">Ranges can have formats applied to individual cells based on conditions.</span></span> <span data-ttu-id="99d34-205">Pour plus d’informations à ce sujet, voir [Appliquer une mise en forme conditionnelle à des plages Excel](excel-add-ins-conditional-formatting.md).</span><span class="sxs-lookup"><span data-stu-id="99d34-205">For more information about this, see [Apply conditional formatting to Excel ranges](excel-add-ins-conditional-formatting.md).</span></span>

## <a name="copy-and-paste"></a><span data-ttu-id="99d34-206">Copier et coller</span><span class="sxs-lookup"><span data-stu-id="99d34-206">Copy and Paste</span></span>

> [!NOTE]
> <span data-ttu-id="99d34-207">Le fonction copyFrom est actuellement disponible uniquement en préversion publique (bêta).</span><span class="sxs-lookup"><span data-stu-id="99d34-207">The copyFrom function is currently available only in public preview (beta).</span></span> <span data-ttu-id="99d34-208">Pour utiliser cette fonctionnalité, vous devez utiliser la bibliothèque bêta du CDN Office.js : https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="99d34-208">To run this sample, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="99d34-209">Si vous utilisez TypeScript ou si votre éditeur de code utilise des fichiers de définition de type TypeScript pour IntelliSense, utilisez https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span><span class="sxs-lookup"><span data-stu-id="99d34-209">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="99d34-210">La fonction copyFrom de la plage reproduit le comportement de copier-coller de l’interface utilisateur Excel.</span><span class="sxs-lookup"><span data-stu-id="99d34-210">Range’s copyFrom function replicates the copy-and-paste behavior of the Excel UI.</span></span> <span data-ttu-id="99d34-211">L’objet plage sur lequel copyFrom est appelé est la destination.</span><span class="sxs-lookup"><span data-stu-id="99d34-211">The range object that copyFrom is called on is the destination.</span></span> <span data-ttu-id="99d34-212">La source à copier est transmise en tant que plage ou qu’adresse de chaîne représentant une plage.</span><span class="sxs-lookup"><span data-stu-id="99d34-212">The source to be copied is passed as a range or a string address representing a range.</span></span> <span data-ttu-id="99d34-213">L’exemple de code suivant copie les données de la plage **A1:E1** dans la plage commençant en **G1** (ce qui aboutit à un collage dans la plage **G1:K1**).</span><span class="sxs-lookup"><span data-stu-id="99d34-213">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range starting at a single cell destination
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="99d34-214">Range.copyFrom a trois paramètres facultatifs.</span><span class="sxs-lookup"><span data-stu-id="99d34-214">Range.copyFrom has three optional parameters.</span></span>

```ts
copyFrom(sourceRange: Range | string, copyType?: "All" | "Formulas" | "Values" | "Formats", skipBlanks?: boolean, transpose?: boolean): void;
``` 

<span data-ttu-id="99d34-215">`copyType` spécifie les données copiées de la source vers la destination.</span><span class="sxs-lookup"><span data-stu-id="99d34-215">`copyType` specifies what data gets copied from the source to the destination.</span></span> 
<span data-ttu-id="99d34-216">`“Formulas”` transfère les formules dans les cellules sources en préservant le positionnement relatif des plages de ces formules.</span><span class="sxs-lookup"><span data-stu-id="99d34-216">`“Formulas”` transfers the formulas in the source cells and preserves the relative positioning of those formulas’ ranges.</span></span> <span data-ttu-id="99d34-217">Les entrées autres que des formules sont copiées telles quelles.</span><span class="sxs-lookup"><span data-stu-id="99d34-217">Any non-formula entries are copied as-is.</span></span> 
<span data-ttu-id="99d34-218">`“Values”` copie les valeurs des données et, s’il s’agit d’une formule, le résultat de celle-ci.</span><span class="sxs-lookup"><span data-stu-id="99d34-218">`“Values”` copies the data values and, in the case of formulas, the result of the formula.</span></span> 
<span data-ttu-id="99d34-219">`“Formats”` copie la mise en forme de la plage, y compris la police, la couleur et d’autres paramètres de mise en forme, mais aucune valeur.</span><span class="sxs-lookup"><span data-stu-id="99d34-219">`“Formats”` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span> 
<span data-ttu-id="99d34-220">`”All”` (option par défaut) copie les données et la mise en forme, en conservant les formules éventuelles des cellules.</span><span class="sxs-lookup"><span data-stu-id="99d34-220">`”All”` (the default option) copies both data and formatting, preserving cells’ formulas if found.</span></span>

<span data-ttu-id="99d34-221">`skipBlanks` définit si les cellules vides sont copiées dans la destination.</span><span class="sxs-lookup"><span data-stu-id="99d34-221">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="99d34-222">Quand la valeur est true, `copyFrom` ignore les cellules vides de la plage source.</span><span class="sxs-lookup"><span data-stu-id="99d34-222">When true, `copyFrom` skips blank cells in the source range.</span></span> <span data-ttu-id="99d34-223">Les cellules ignorées ne remplacent pas les données existantes dans les cellules correspondantes de la plage de destination.</span><span class="sxs-lookup"><span data-stu-id="99d34-223">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="99d34-224">La valeur par défaut est false.</span><span class="sxs-lookup"><span data-stu-id="99d34-224">The default is False.</span></span>

<span data-ttu-id="99d34-225">L’exemple de code et les images suivants illustrent ce comportement dans un scénario simple.</span><span class="sxs-lookup"><span data-stu-id="99d34-225">The following code sample and images demonstrate this behavior in a simple scenario.</span></span> 

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

<span data-ttu-id="99d34-226">*Avant exécution de la fonction précédente.*</span><span class="sxs-lookup"><span data-stu-id="99d34-226">*Before the preceeding function has been run.*</span></span>

![Données dans Excel avant exécution de la méthode de copie de la plage.](../images/excel-range-copyfrom-skipblanks-before.png)

<span data-ttu-id="99d34-228">*Après exécution de la fonction précédente.*</span><span class="sxs-lookup"><span data-stu-id="99d34-228">*After the preceeding function has been run.*</span></span>

![Données dans Excel après exécution de la méthode de copie de la plage.](../images/excel-range-copyfrom-skipblanks-after.png)

<span data-ttu-id="99d34-230">`transpose` détermine si les données sont ou non transposées, ce qui signifie que ses lignes et colonnes sont permutées dans l’emplacement source.</span><span class="sxs-lookup"><span data-stu-id="99d34-230">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span> <span data-ttu-id="99d34-231">Une plage transposée est renversée le long de la diagonale principale, de sorte que les lignes **1**, **2** et **3** deviennent les colonnes **A**, **B** et **C**.</span><span class="sxs-lookup"><span data-stu-id="99d34-231">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span> 


## <a name="see-also"></a><span data-ttu-id="99d34-232">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="99d34-232">See also</span></span>

- [<span data-ttu-id="99d34-233">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="99d34-233">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)

