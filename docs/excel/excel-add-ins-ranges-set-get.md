---
title: Définir et obtenir la plage sélectionnée à l’aide de l Excel API JavaScript
description: Découvrez comment utiliser l’API JavaScript Excel pour définir et obtenir la plage sélectionnée à l’aide de Excel API JavaScript.
ms.date: 07/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 623ba5c1b9e76151d4a2c4b169e655236b37e8c8
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290781"
---
# <a name="set-and-get-the-selected-range-using-the-excel-javascript-api"></a><span data-ttu-id="c53b5-103">Définir et obtenir la plage sélectionnée à l’aide de l Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="c53b5-103">Set and get the selected range using the Excel JavaScript API</span></span>

<span data-ttu-id="c53b5-104">Cet article fournit des exemples de code qui définissent et obtiennent la plage sélectionnée avec Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="c53b5-104">This article provides code samples that set and get the selected range with the Excel JavaScript API.</span></span> <span data-ttu-id="c53b5-105">Pour obtenir la liste complète des propriétés et méthodes que l’objet prend en `Range` charge, [voir Excel. Classe Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="c53b5-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a><span data-ttu-id="c53b5-106">Définir la plage sélectionnée</span><span class="sxs-lookup"><span data-stu-id="c53b5-106">Set the selected range</span></span>

<span data-ttu-id="c53b5-107">L’exemple de code suivant sélectionne la plage **B2:E6** dans la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="c53b5-107">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a><span data-ttu-id="c53b5-108">Plage sélectionnée  B2:E6</span><span class="sxs-lookup"><span data-stu-id="c53b5-108">Selected range B2:E6</span></span>

![Plage sélectionnée en Excel.](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="c53b5-110">Obtenir la plage sélectionnée</span><span class="sxs-lookup"><span data-stu-id="c53b5-110">Get the selected range</span></span>

<span data-ttu-id="c53b5-111">L’exemple de code suivant obtient la plage sélectionnée, charge sa propriété et écrit `address` un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="c53b5-111">The following code sample gets the selected range, loads its `address` property, and writes a message to the console.</span></span>

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

## <a name="select-the-edge-of-a-used-range"></a><span data-ttu-id="c53b5-112">Sélectionner le bord d’une plage utilisée</span><span class="sxs-lookup"><span data-stu-id="c53b5-112">Select the edge of a used range</span></span>

<span data-ttu-id="c53b5-113">Les méthodes [Range.getRangeEdge](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_) et [Range.getExtendedRange](/javascript/api/excel/excel.range#getExtendedRange_directionString__activeCell_) vous permet de répliquer le comportement des raccourcis de sélection du clavier, en sélectionnant le bord de la plage utilisée en fonction de la plage actuellement sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="c53b5-113">The [Range.getRangeEdge](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_) and [Range.getExtendedRange](/javascript/api/excel/excel.range#getExtendedRange_directionString__activeCell_) methods let your add-in replicate the behavior of the keyboard selection shortcuts, selecting the edge of the used range based on the currently selected range.</span></span> <span data-ttu-id="c53b5-114">Pour en savoir plus sur les plages utilisées, voir [Obtenir une plage utilisée.](excel-add-ins-ranges-get.md#get-used-range)</span><span class="sxs-lookup"><span data-stu-id="c53b5-114">To learn more about used ranges, see [Get used range](excel-add-ins-ranges-get.md#get-used-range).</span></span>

<span data-ttu-id="c53b5-115">Dans la capture d’écran suivante, la plage utilisée est le tableau avec des valeurs dans chaque cellule, **C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="c53b5-115">In the following screenshot, the used range is the table with values in each cell, **C5:F12**.</span></span> <span data-ttu-id="c53b5-116">Les cellules vides en dehors de ce tableau sont en dehors de la plage utilisée.</span><span class="sxs-lookup"><span data-stu-id="c53b5-116">The empty cells outside this table are outside the used range.</span></span>

![Tableau avec des données de C5:F12 Excel.](../images/excel-ranges-used-range.png)

### <a name="select-the-cell-at-the-edge-of-the-current-used-range"></a><span data-ttu-id="c53b5-118">Sélectionner la cellule au bord de la plage utilisée actuelle</span><span class="sxs-lookup"><span data-stu-id="c53b5-118">Select the cell at the edge of the current used range</span></span>

<span data-ttu-id="c53b5-119">L’exemple de code suivant montre comment utiliser la méthode pour sélectionner la cellule au bord le plus proche de la plage utilisée actuelle, dans `Range.getRangeEdge` la direction vers le haut.</span><span class="sxs-lookup"><span data-stu-id="c53b5-119">The following code sample shows how use the `Range.getRangeEdge` method to select the cell at the furthest edge of the current used range, in the direction up.</span></span> <span data-ttu-id="c53b5-120">Cette action correspond au résultat de l’utilisation du raccourci clavier de touche fléchée Ctrl+Haut pendant qu’une plage est sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="c53b5-120">This action matches the result of using the Ctrl+Up arrow key keyboard shortcut while a range is selected.</span></span>

```js
Excel.run(function (context) {
    // Get the selected range.
    var range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    var direction = Excel.KeyboardDirection.up;

    // Get the active cell in the workbook.
    var activeCell = context.workbook.getActiveCell();

    // Get the top-most cell of the current used range.
    // This method acts like the Ctrl+Up arrow key keyboard shortcut while a range is selected.
    var rangeEdge = range.getRangeEdge(
      direction,
      activeCell
    );
    rangeEdge.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="before-selecting-the-cell-at-the-edge-of-the-used-range"></a><span data-ttu-id="c53b5-121">Avant de sélectionner la cellule au bord de la plage utilisée</span><span class="sxs-lookup"><span data-stu-id="c53b5-121">Before selecting the cell at the edge of the used range</span></span>

<span data-ttu-id="c53b5-122">La capture d’écran suivante montre une plage utilisée et une plage sélectionnée dans la plage utilisée.</span><span class="sxs-lookup"><span data-stu-id="c53b5-122">The following screenshot shows a used range and a selected range within the used range.</span></span> <span data-ttu-id="c53b5-123">La plage utilisée est un tableau avec des données **au niveau de C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="c53b5-123">The used range is a table with data at **C5:F12**.</span></span> <span data-ttu-id="c53b5-124">Dans ce tableau, la plage **D8:E9** est sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="c53b5-124">Inside this table, the range **D8:E9** is selected.</span></span> <span data-ttu-id="c53b5-125">Cette sélection est à *l’état* antérieur, avant l’exécution de la `Range.getRangeEdge` méthode.</span><span class="sxs-lookup"><span data-stu-id="c53b5-125">This selection is the *before* state, prior to running the `Range.getRangeEdge` method.</span></span>

![Tableau avec des données de C5:F12 Excel.](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-the-cell-at-the-edge-of-the-used-range"></a><span data-ttu-id="c53b5-128">Après avoir sélectionné la cellule au bord de la plage utilisée</span><span class="sxs-lookup"><span data-stu-id="c53b5-128">After selecting the cell at the edge of the used range</span></span>

<span data-ttu-id="c53b5-129">La capture d’écran suivante montre le même tableau que la capture d’écran précédente, avec des données de la plage **C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="c53b5-129">The following screenshot shows the same table as the preceding screenshot, with data in the range **C5:F12**.</span></span> <span data-ttu-id="c53b5-130">Dans ce tableau, la plage **D5** est sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="c53b5-130">Inside this table, the range **D5** is selected.</span></span> <span data-ttu-id="c53b5-131">Cette sélection *s’exécute après* l’exécution de la méthode pour sélectionner la cellule au bord de la plage utilisée dans la direction vers le `Range.getRangeEdge` haut.</span><span class="sxs-lookup"><span data-stu-id="c53b5-131">This selection is *after* state, after running the `Range.getRangeEdge` method to select the cell at the edge of the used range in the up direction.</span></span>

![Tableau avec des données de C5:F12 Excel.](../images/excel-ranges-used-range-d5.png)

### <a name="select-all-cells-from-current-range-to-furthest-edge-of-used-range"></a><span data-ttu-id="c53b5-134">Sélectionner toutes les cellules de la plage actuelle au bord le plus proche de la plage utilisée</span><span class="sxs-lookup"><span data-stu-id="c53b5-134">Select all cells from current range to furthest edge of used range</span></span>

<span data-ttu-id="c53b5-135">L’exemple de code suivant montre comment utiliser la méthode pour sélectionner toutes les cellules de la plage actuellement sélectionnée au bord le plus proche de la plage utilisée, dans la direction vers le `Range.getExtendedRange` bas.</span><span class="sxs-lookup"><span data-stu-id="c53b5-135">The following code sample shows how use the `Range.getExtendedRange` method to to select all the cells from the currently selected range to the furthest edge of the used range, in the direction down.</span></span> <span data-ttu-id="c53b5-136">Cette action correspond au résultat de l’utilisation du raccourci clavier avec touches de direction Ctrl+Shift+Bas pendant qu’une plage est sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="c53b5-136">This action matches the result of using the Ctrl+Shift+Down arrow key keyboard shortcut while a range is selected.</span></span>

```js
Excel.run(function (context) {
    // Get the selected range.
    var range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    var direction = Excel.KeyboardDirection.down;

    // Get the active cell in the workbook.
    var activeCell = context.workbook.getActiveCell();

    // Get all the cells from the currently selected range to the bottom-most edge of the used range.
    // This method acts like the Ctrl+Shift+Down arrow key keyboard shortcut while a range is selected.
    var extendedRange = range.getExtendedRange(
      direction,
      activeCell
    );
    extendedRange.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="before-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a><span data-ttu-id="c53b5-137">Avant de sélectionner toutes les cellules de la plage actuelle au bord de la plage utilisée</span><span class="sxs-lookup"><span data-stu-id="c53b5-137">Before selecting all the cells from the current range to the edge of the used range</span></span>

<span data-ttu-id="c53b5-138">La capture d’écran suivante montre une plage utilisée et une plage sélectionnée dans la plage utilisée.</span><span class="sxs-lookup"><span data-stu-id="c53b5-138">The following screenshot shows a used range and a selected range within the used range.</span></span> <span data-ttu-id="c53b5-139">La plage utilisée est un tableau avec des données **au niveau de C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="c53b5-139">The used range is a table with data at **C5:F12**.</span></span> <span data-ttu-id="c53b5-140">Dans ce tableau, la plage **D8:E9** est sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="c53b5-140">Inside this table, the range **D8:E9** is selected.</span></span> <span data-ttu-id="c53b5-141">Cette sélection est à *l’état* antérieur, avant l’exécution de la `Range.getExtendedRange` méthode.</span><span class="sxs-lookup"><span data-stu-id="c53b5-141">This selection is the *before* state, prior to running the `Range.getExtendedRange` method.</span></span>

![Tableau avec des données de C5:F12 Excel.](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a><span data-ttu-id="c53b5-144">Après avoir sélectionné toutes les cellules de la plage actuelle au bord de la plage utilisée</span><span class="sxs-lookup"><span data-stu-id="c53b5-144">After selecting all the cells from the current range to the edge of the used range</span></span>

<span data-ttu-id="c53b5-145">La capture d’écran suivante montre le même tableau que la capture d’écran précédente, avec des données de la plage **C5:F12**.</span><span class="sxs-lookup"><span data-stu-id="c53b5-145">The following screenshot shows the same table as the preceding screenshot, with data in the range **C5:F12**.</span></span> <span data-ttu-id="c53b5-146">Dans ce tableau, la plage **D8:E12** est sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="c53b5-146">Inside this table, the range **D8:E12** is selected.</span></span> <span data-ttu-id="c53b5-147">Cette sélection *s’exécute* après l’exécution de la méthode pour sélectionner toutes les cellules de la plage actuelle au bord de la plage utilisée dans `Range.getExtendedRange` la direction vers le bas.</span><span class="sxs-lookup"><span data-stu-id="c53b5-147">This selection is *after* state, after running the `Range.getExtendedRange` method to select all the cells from the current range to the edge of the used range in the down direction.</span></span>

![Tableau avec des données de C5:F12 Excel.](../images/excel-ranges-used-range-d8-e12.png)

## <a name="see-also"></a><span data-ttu-id="c53b5-150">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c53b5-150">See also</span></span>

- [<span data-ttu-id="c53b5-151">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="c53b5-151">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="c53b5-152">Utiliser des cellules à l’aide Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="c53b5-152">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="c53b5-153">Définir et obtenir des valeurs de plage, du texte ou des formules à l’aide Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="c53b5-153">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
- [<span data-ttu-id="c53b5-154">Définir le format de plage à l’aide Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="c53b5-154">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)
