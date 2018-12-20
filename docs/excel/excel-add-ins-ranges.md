---
title: Utilisation de plages à l’aide de l’API JavaScript pour Excel (fondamental)
description: ''
ms.date: 12/14/2018
ms.openlocfilehash: 4c64abec1f79bd1194a106e46b8a6fe6c4b71d07
ms.sourcegitcommit: 09f124fac7b2e711e1a8be562a99624627c0699e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/15/2018
ms.locfileid: "27283101"
---
# <a name="work-with-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="c2a81-102">Utilisation de plages à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="c2a81-102">Work with ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="c2a81-103">Cet article fournit des exemples de code qui expliquent comment effectuer des tâches courantes avec des plages à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="c2a81-103">This article provides code samples that show how to perform common tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="c2a81-104">Pour obtenir une liste complète des propriétés et des méthodes prises en charge par l’objet **Range**, reportez-vous à la rubrique [Objet Range (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="c2a81-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span></span>

> [!NOTE]
> <span data-ttu-id="c2a81-105">Pour plus d’exemples de code qui montrent comment effectuer des tâches plus avancées avec des plages, consultez l’article [Utiliser les plages à l’aide de l’API JavaScript Excel (avancé)](excel-add-ins-ranges-advanced.md).</span><span class="sxs-lookup"><span data-stu-id="c2a81-105">For code samples that show how to perform more advanced tasks with ranges, see [Work with ranges using the Excel JavaScript API (advanced)](excel-add-ins-ranges-advanced.md).</span></span>

## <a name="get-a-range"></a><span data-ttu-id="c2a81-106">Obtenir une plage</span><span class="sxs-lookup"><span data-stu-id="c2a81-106">Get a range</span></span>

<span data-ttu-id="c2a81-107">Les exemples suivants montrent les différentes façons d’obtenir une référence à une plage dans une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="c2a81-107">The following examples show different ways to get a reference to a range within a worksheet.</span></span>

### <a name="get-range-by-address"></a><span data-ttu-id="c2a81-108">Obtenir une plage en fonction d’une adresse</span><span class="sxs-lookup"><span data-stu-id="c2a81-108">Get range by address</span></span>

<span data-ttu-id="c2a81-109">L’exemple de code suivant obtient la plage ayant l’adresse **B2 : B5** à partir de la feuille de calcul nommée **Sample**, charge sa propriété **address** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="c2a81-109">The following code sample gets the range with address **B2:B5** from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

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

### <a name="get-range-by-name"></a><span data-ttu-id="c2a81-110">Obtenir une plage en fonction d’un nom</span><span class="sxs-lookup"><span data-stu-id="c2a81-110">Get range by name</span></span>

<span data-ttu-id="c2a81-111">L’exemple de code suivant obtient la plage nommée **MyRange** à partir de la feuille de calcul nommée **Sample**, charge sa propriété **address** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="c2a81-111">The following code sample gets the range named **MyRange** from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

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

### <a name="get-used-range"></a><span data-ttu-id="c2a81-112">Obtenir une plage utilisée</span><span class="sxs-lookup"><span data-stu-id="c2a81-112">Get used range</span></span>

<span data-ttu-id="c2a81-113">L’exemple de code suivant obtient la plage  utilisée dans la feuille de calcul nommée **Sample**, charge sa propriété **address** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="c2a81-113">The following code sample gets the used range from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span> <span data-ttu-id="c2a81-114">La plage utilisée est la plus petite plage qui englobe toutes les cellules de la feuille de calcul auxquelles une valeur ou un format est affecté.</span><span class="sxs-lookup"><span data-stu-id="c2a81-114">The used range is the smallest range that encompasses any cells in the worksheet that have a value or formatting assigned to them.</span></span> <span data-ttu-id="c2a81-115">Si la feuille de calcul entière est vide, la méthode **getUsedRange()** renvoie une plage qui se compose d’uniquement de la cellule en haut à gauche de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="c2a81-115">If the entire worksheet is blank, the **getUsedRange()** method returns a range that consists of only the top-left cell in the worksheet.</span></span>

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

### <a name="get-entire-range"></a><span data-ttu-id="c2a81-116">Obtenir l’intégralité d’une plage</span><span class="sxs-lookup"><span data-stu-id="c2a81-116">Get entire range</span></span>

<span data-ttu-id="c2a81-117">L’exemple de code suivant obtient l’intégralité de la plage de la feuille de calcul à partir de la feuille de calcul nommée **Sample**, charge sa propriété **address** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="c2a81-117">The following code sample gets the entire worksheet range from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

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

## <a name="insert-a-range-of-cells"></a><span data-ttu-id="c2a81-118">Insérer une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="c2a81-118">Insert a range of cells</span></span>

<span data-ttu-id="c2a81-119">L’exemple de code suivant insère une plage de cellules dans l’emplacement **B4:E4** et déplace les autres cellules vers le bas pour laisser de l’espace pour les nouvelles cellules.</span><span class="sxs-lookup"><span data-stu-id="c2a81-119">The following code sample inserts a range of cells in location **B4:E4** and shifts other cells down to provide space for the new cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);
    
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="c2a81-120">**Données avant l’insertion de la plage**</span><span class="sxs-lookup"><span data-stu-id="c2a81-120">**Data before range is inserted**</span></span>

![Données dans Excel avant l’insertion de la plage](../images/excel-ranges-start.png)

<span data-ttu-id="c2a81-122">**Données après l’insertion de la plage**</span><span class="sxs-lookup"><span data-stu-id="c2a81-122">**Data after range is inserted**</span></span>

![Données dans Excel après l’insertion de plage](../images/excel-ranges-after-insert.png)

## <a name="clear-a-range-of-cells"></a><span data-ttu-id="c2a81-124">Effacer une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="c2a81-124">Clear a range of cells</span></span>

<span data-ttu-id="c2a81-125">L’exemple de code suivant efface tout le contenu et la mise en forme des cellules de la plage **E2 : E5**.</span><span class="sxs-lookup"><span data-stu-id="c2a81-125">The following code sample clears all contents and formatting of cells in the range **E2:E5**.</span></span>  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="c2a81-126">**Données avant l’effacement de la plage**</span><span class="sxs-lookup"><span data-stu-id="c2a81-126">**Data before range is cleared**</span></span>

![Données dans Excel avant l’effacement de la plage](../images/excel-ranges-start.png)

<span data-ttu-id="c2a81-128">**Données après l’effacement de plage**</span><span class="sxs-lookup"><span data-stu-id="c2a81-128">**Data after range is cleared**</span></span>

![Données dans Excel après l’effacement de plage](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a><span data-ttu-id="c2a81-130">Supprimer une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="c2a81-130">Delete a range of cells</span></span>

<span data-ttu-id="c2a81-131">L’exemple de code suivant supprime les cellules de la plage **B4:E4** et déplace les autres cellules vers le haut pour remplir l’espace libre suite à la suppression des cellules.</span><span class="sxs-lookup"><span data-stu-id="c2a81-131">The following code sample deletes the cells in the range **B4:E4** and shift other cells up to fill the space that was vacated by the deleted cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="c2a81-132">**Données avant la suppression d’une plage**</span><span class="sxs-lookup"><span data-stu-id="c2a81-132">**Data before range is deleted**</span></span>

![Données dans Excel avant la suppression d’une plage](../images/excel-ranges-start.png)

<span data-ttu-id="c2a81-134">**Données après la suppression d’une plage**</span><span class="sxs-lookup"><span data-stu-id="c2a81-134">**Data after range is deleted**</span></span>

![Données dans Excel après la suppression d’une plage](../images/excel-ranges-after-delete.png)

## <a name="set-the-selected-range"></a><span data-ttu-id="c2a81-136">Définir la plage sélectionnée</span><span class="sxs-lookup"><span data-stu-id="c2a81-136">Set the selected range</span></span>

<span data-ttu-id="c2a81-137">L’exemple de code suivant sélectionne la plage **B2:E6** dans la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="c2a81-137">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="c2a81-138">**Plage sélectionnée  B2:E6**</span><span class="sxs-lookup"><span data-stu-id="c2a81-138">**Selected range B2:E6**</span></span>

![Plage sélectionnée dans Excel](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="c2a81-140">Obtenir la plage sélectionnée</span><span class="sxs-lookup"><span data-stu-id="c2a81-140">Get the selected range</span></span>

<span data-ttu-id="c2a81-141">L’exemple de code suivant recherche la plage  sélectionnée, charge sa propriété **address** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="c2a81-141">The following code sample gets the selected range, loads its **address** property, and writes a message to the console.</span></span> 

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

## <a name="set-values-or-formulas"></a><span data-ttu-id="c2a81-142">Définir des valeurs ou des formules</span><span class="sxs-lookup"><span data-stu-id="c2a81-142">Set values or formulas</span></span>

<span data-ttu-id="c2a81-143">Les exemples suivants indiquent comment définir des valeurs et des formules pour une cellule unique ou une plage de cellules.</span><span class="sxs-lookup"><span data-stu-id="c2a81-143">The following examples show how to set values and formulas for a single cell or a range of cells.</span></span>

### <a name="set-value-for-a-single-cell"></a><span data-ttu-id="c2a81-144">Définir une valeur pour une cellule unique</span><span class="sxs-lookup"><span data-stu-id="c2a81-144">Set value for a single cell</span></span>

<span data-ttu-id="c2a81-145">L’exemple de code suivant définit la valeur de la cellule **C3** sur « 5 », puis définit la largeur des colonnes pour mieux s’adapter aux données.</span><span class="sxs-lookup"><span data-stu-id="c2a81-145">The following code sample sets the value of cell **C3** to "5" and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="c2a81-146">**Données avant la mise à jour de la valeur de la cellule**</span><span class="sxs-lookup"><span data-stu-id="c2a81-146">**Data before cell value is updated**</span></span>

![Données dans Excel avant la mise à jour de la valeur de la cellule](../images/excel-ranges-set-start.png)

<span data-ttu-id="c2a81-148">**Données après la mise à jour de la valeur de la cellule**</span><span class="sxs-lookup"><span data-stu-id="c2a81-148">**Data after cell value is updated**</span></span>

![Données dans Excel après la mise à jour de la valeur de la cellule](../images/excel-ranges-set-cell-value.png)

### <a name="set-values-for-a-range-of-cells"></a><span data-ttu-id="c2a81-150">Définir des valeurs pour une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="c2a81-150">Set values for a range of cells</span></span>

<span data-ttu-id="c2a81-151">L’exemple de code suivant définit les valeurs des cellules de la plage **B5:D5**, puis définit la largeur des colonnes pour mieux s’adapter aux données.</span><span class="sxs-lookup"><span data-stu-id="c2a81-151">The following code sample sets values for the cells in the range **B5:D5** and then sets the width of the columns to best fit the data.</span></span>

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

<span data-ttu-id="c2a81-152">**Données avant la mise à jour des valeurs des cellules**</span><span class="sxs-lookup"><span data-stu-id="c2a81-152">**Data before cell values are updated**</span></span>

![Données dans Excel avant la mise à jour des valeurs des cellules](../images/excel-ranges-set-start.png)

<span data-ttu-id="c2a81-154">**Données après la mise à jour des valeurs des cellules**</span><span class="sxs-lookup"><span data-stu-id="c2a81-154">**Data after cell values are updated**</span></span>

![Données dans Excel après la mise à jour des valeurs des cellules](../images/excel-ranges-set-cell-values.png)

### <a name="set-formula-for-a-single-cell"></a><span data-ttu-id="c2a81-156">Définir la formule d’une cellule unique</span><span class="sxs-lookup"><span data-stu-id="c2a81-156">Set formula for a single cell</span></span>

<span data-ttu-id="c2a81-157">L’exemple de code suivant définit une formule pour la cellule **E3**, puis définit la largeur des colonnes pour mieux s’adapter aux données.</span><span class="sxs-lookup"><span data-stu-id="c2a81-157">The following code sample sets a formula for cell **E3** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="c2a81-158">**Données avant la définition de la formule de la cellule**</span><span class="sxs-lookup"><span data-stu-id="c2a81-158">**Data before cell formula is set**</span></span>

![Données dans Excel avant la définition de la formule de la cellule](../images/excel-ranges-start-set-formula.png)

<span data-ttu-id="c2a81-160">**Données après la définition de la formule de la cellule**</span><span class="sxs-lookup"><span data-stu-id="c2a81-160">**Data after cell formula is set**</span></span>

![Données dans Excel après la définition de la formule de la cellule](../images/excel-ranges-set-formula.png)

### <a name="set-formulas-for-a-range-of-cells"></a><span data-ttu-id="c2a81-162">Définir des formules pour une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="c2a81-162">Set formulas for a range of cells</span></span>

<span data-ttu-id="c2a81-163">L’exemple de code ci-dessous définit des formules pour les cellules de la plage **E2:E6**, puis définit la largeur des colonnes pour mieux s’adapter aux données.</span><span class="sxs-lookup"><span data-stu-id="c2a81-163">The following code sample sets formulas for cells in the range **E2:E6** and then sets the width of the columns to best fit the data.</span></span>

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

<span data-ttu-id="c2a81-164">**Données avant la définition des formules des cellules**</span><span class="sxs-lookup"><span data-stu-id="c2a81-164">**Data before cell formulas are set**</span></span>

![Données dans Excel avant la définition des formules des cellules](../images/excel-ranges-start-set-formula.png)

<span data-ttu-id="c2a81-166">**Données après la définition des formules des cellules**</span><span class="sxs-lookup"><span data-stu-id="c2a81-166">**Data after cell formulas are set**</span></span>

![Données dans Excel après la définition des formules des cellules](../images/excel-ranges-set-formulas.png)

## <a name="get-values-text-or-formulas"></a><span data-ttu-id="c2a81-168">Obtenir des valeurs, du texte ou des formules</span><span class="sxs-lookup"><span data-stu-id="c2a81-168">Get values, text, or formulas</span></span>

<span data-ttu-id="c2a81-169">Ces exemples montrent comment obtenir des valeurs, du texte et des formules à partir d’une plage de cellules.</span><span class="sxs-lookup"><span data-stu-id="c2a81-169">These examples show how to get values, text, and formulas from a range of cells.</span></span>

### <a name="get-values-from-a-range-of-cells"></a><span data-ttu-id="c2a81-170">Obtenir des valeurs à partir d’une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="c2a81-170">Get values from a range of cells</span></span>

<span data-ttu-id="c2a81-171">L’exemple de code suivant obtient la plage **B2:E6**, charge la propriété **values** et écrit les valeurs dans la console.</span><span class="sxs-lookup"><span data-stu-id="c2a81-171">The following code sample gets the range **B2:E6**, loads its **values** property, and writes the values to the console.</span></span> <span data-ttu-id="c2a81-172">La propriété **values** d’une plage spécifie les valeurs brutes contenues dans les cellules.</span><span class="sxs-lookup"><span data-stu-id="c2a81-172">The **values** property of a range specifies the raw values that the cells contain.</span></span> <span data-ttu-id="c2a81-173">Même si certaines cellules d’une plage contiennent des formules, la propriété **values** de la plage spécifie les valeurs brutes des cellules, et non des formules.</span><span class="sxs-lookup"><span data-stu-id="c2a81-173">Even if some cells in a range contain formulas, the **values** property of the range specifies the raw values for those cells, not any of the formulas.</span></span>

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

<span data-ttu-id="c2a81-174">**Données de la plage (les valeurs de la colonne E sont le résultat des formules)**</span><span class="sxs-lookup"><span data-stu-id="c2a81-174">**Data in range (values in column E are a result of formulas)**</span></span>

![Données dans Excel après la définition des formules des cellules](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="c2a81-176">**range.values (comme consigné dans la console par l’exemple de code ci-dessus)**</span><span class="sxs-lookup"><span data-stu-id="c2a81-176">**range.values (as logged to the console by the code sample above)**</span></span>

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

### <a name="get-text-from-a-range-of-cells"></a><span data-ttu-id="c2a81-177">Obtenir du texte à partir d’une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="c2a81-177">Get text from a range of cells</span></span>

<span data-ttu-id="c2a81-178">L’exemple de code suivant obtient la plage **B2:E6**, charge sa propriété **text** et l’écrit dans la console.</span><span class="sxs-lookup"><span data-stu-id="c2a81-178">The following code sample gets the range **B2:E6**, loads its **text** property, and writes it to the console.</span></span>  <span data-ttu-id="c2a81-179">La propriété **text** d’une plage spécifie les valeurs d’affichage pour les cellules de la plage.</span><span class="sxs-lookup"><span data-stu-id="c2a81-179">The **text** property of a range specifies the display values for cells in the range.</span></span> <span data-ttu-id="c2a81-180">Même si certaines cellules d’une plage contiennent des formules, la propriété **text** de la plage indique les valeurs d’affichage pour ces cellules, et non des formules.</span><span class="sxs-lookup"><span data-stu-id="c2a81-180">Even if some cells in a range contain formulas, the **text** property of the range specifies the display values for those cells, not any of the formulas.</span></span>

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

<span data-ttu-id="c2a81-181">**Données de la plage (les valeurs de la colonne E sont le résultat des formules)**</span><span class="sxs-lookup"><span data-stu-id="c2a81-181">**Data in range (values in column E are a result of formulas)**</span></span>

![Données dans Excel après la définition des formules des cellules](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="c2a81-183">**range.text (comme consigné dans la console par l’exemple de code ci-dessus)**</span><span class="sxs-lookup"><span data-stu-id="c2a81-183">**range.text (as logged to the console by the code sample above)**</span></span>

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

### <a name="get-formulas-from-a-range-of-cells"></a><span data-ttu-id="c2a81-184">Obtenir des formules à partir d’une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="c2a81-184">Get formulas from a range of cells</span></span>

<span data-ttu-id="c2a81-185">L’exemple de code suivant obtient la plage **B2:E6**, charge sa propriété **formulas** et l’écrit dans la console.</span><span class="sxs-lookup"><span data-stu-id="c2a81-185">The following code sample gets the range **B2:E6**, loads its **formulas** property, and writes it to the console.</span></span>  <span data-ttu-id="c2a81-186">La propriété **formulas** d’une plage spécifie les formules pour les cellules de la plage contenant des formules et des valeurs brutes pour les cellules de la plage ne contenant pas de formule.</span><span class="sxs-lookup"><span data-stu-id="c2a81-186">The **formulas** property of a range specifies the formulas for cells in the range that contain formulas and the raw values for cells in the range that do not contain formulas.</span></span>

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

<span data-ttu-id="c2a81-187">**Données de la plage (les valeurs de la colonne E sont le résultat des formules)**</span><span class="sxs-lookup"><span data-stu-id="c2a81-187">**Data in range (values in column E are a result of formulas)**</span></span>

![Données dans Excel après la définition des formules des cellules](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="c2a81-189">**range.formulas (comme consigné dans la console par l’exemple de code ci-dessus)**</span><span class="sxs-lookup"><span data-stu-id="c2a81-189">**range.formulas (as logged to the console by the code sample above)**</span></span>

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

## <a name="set-range-format"></a><span data-ttu-id="c2a81-190">Définir le format de plage</span><span class="sxs-lookup"><span data-stu-id="c2a81-190">Set range format</span></span>

<span data-ttu-id="c2a81-191">Les exemples ci-dessous indiquent comment définir la couleur de police, la couleur de remplissage et le format de nombre pour des cellules dans une plage.</span><span class="sxs-lookup"><span data-stu-id="c2a81-191">The following examples show how to set font color, fill color, and number format for cells in a range.</span></span>

### <a name="set-font-color-and-fill-color"></a><span data-ttu-id="c2a81-192">Définir la couleur de police et la couleur de remplissage</span><span class="sxs-lookup"><span data-stu-id="c2a81-192">Set font color and fill color</span></span>

<span data-ttu-id="c2a81-193">L’exemple de code ci-dessous définit la couleur de police et la couleur de remplissage des cellules de la plage **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="c2a81-193">The following code sample sets the font color and fill color for cells in range **B2:E2**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";;
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="c2a81-194">**Données de la plage avant la définition de la couleur de police et de la couleur de remplissage**</span><span class="sxs-lookup"><span data-stu-id="c2a81-194">**Data in range before font color and fill color are set**</span></span>

![Données dans Excel de la plage avant la définition de la couleur de police et de la couleur de remplissage](../images/excel-ranges-format-before.png)

<span data-ttu-id="c2a81-196">**Données de la plage après la définition de la couleur de police et de la couleur de remplissage**</span><span class="sxs-lookup"><span data-stu-id="c2a81-196">**Data in range after font color and fill color are set**</span></span>

![Données dans Excel après la définition du format](../images/excel-ranges-format-font-and-fill.png)

### <a name="set-number-format"></a><span data-ttu-id="c2a81-198">Définir le format de nombre</span><span class="sxs-lookup"><span data-stu-id="c2a81-198">Set number format</span></span>

<span data-ttu-id="c2a81-199">L’exemple de code ci-dessous définit le format de nombre des cellules dans la plage **D3:E5**.</span><span class="sxs-lookup"><span data-stu-id="c2a81-199">The following code sample sets the number format for the cells in range **D3:E5**.</span></span>

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

<span data-ttu-id="c2a81-200">**Données de la plage avant la définition du format de nombre**</span><span class="sxs-lookup"><span data-stu-id="c2a81-200">**Data in range before number format is set**</span></span>

![Données dans Excel de la plage avant la définition de la couleur de police et de la couleur de remplissage](../images/excel-ranges-format-font-and-fill.png)

<span data-ttu-id="c2a81-202">**Données de la plage après la définition du format de nombre**</span><span class="sxs-lookup"><span data-stu-id="c2a81-202">**Data in range after number format is set**</span></span>

![Données dans Excel après la définition du format](../images/excel-ranges-format-numbers.png)

### <a name="conditional-formatting-of-ranges"></a><span data-ttu-id="c2a81-204">Mise en forme conditionnelle de plages</span><span class="sxs-lookup"><span data-stu-id="c2a81-204">Conditional formatting of ranges</span></span>

<span data-ttu-id="c2a81-205">Des plages peuvent présenter une mise en forme de cellules individuelles en fonction de certaines conditions.</span><span class="sxs-lookup"><span data-stu-id="c2a81-205">Ranges can have formats applied to individual cells based on conditions.</span></span> <span data-ttu-id="c2a81-206">Pour plus d’informations à ce sujet, consultez l’article [Appliquer une mise en forme conditionnelle à des plages Excel](excel-add-ins-conditional-formatting.md).</span><span class="sxs-lookup"><span data-stu-id="c2a81-206">For more information about this, see [Apply conditional formatting to Excel ranges](excel-add-ins-conditional-formatting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="c2a81-207">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c2a81-207">See also</span></span>

- [<span data-ttu-id="c2a81-208">Utiliser les plages à l’aide de l’API JavaScript Excel (avancé)</span><span class="sxs-lookup"><span data-stu-id="c2a81-208">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-advanced.md)
- [<span data-ttu-id="c2a81-209">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="c2a81-209">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)