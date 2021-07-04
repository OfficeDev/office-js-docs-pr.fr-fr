---
title: Utiliser des antécédents et des dépendances de formule à l’aide Excel API JavaScript
description: Découvrez comment utiliser l’API JavaScript Excel pour récupérer les antécédents et les dépendances de formule.
ms.date: 07/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: bf92400af00df42ac245b9a2d3ff5e72512b5722
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290774"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a><span data-ttu-id="0399f-103">Obtenir des antécédents et des dépendances de formule à l’aide Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="0399f-103">Get formula precedents and dependents using the Excel JavaScript API</span></span>

<span data-ttu-id="0399f-104">Excel formules font souvent référence à d’autres cellules.</span><span class="sxs-lookup"><span data-stu-id="0399f-104">Excel formulas often refer to other cells.</span></span> <span data-ttu-id="0399f-105">Ces références entre cellules sont appelées « antécédents » et « dépendants ».</span><span class="sxs-lookup"><span data-stu-id="0399f-105">These cross-cell references are known as "precedents" and "dependents".</span></span> <span data-ttu-id="0399f-106">Un précédent est une cellule qui fournit des données à une formule.</span><span class="sxs-lookup"><span data-stu-id="0399f-106">A precedent is a cell that provides data to a formula.</span></span> <span data-ttu-id="0399f-107">Une cellule dépendante est une cellule qui contient une formule qui fait référence à d’autres cellules.</span><span class="sxs-lookup"><span data-stu-id="0399f-107">A dependent is a cell that contains a formula that refers to other cells.</span></span> <span data-ttu-id="0399f-108">Pour en savoir plus sur Excel fonctionnalités liées aux relations entre les cellules, voir Afficher les relations entre les [formules et les cellules.](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507)</span><span class="sxs-lookup"><span data-stu-id="0399f-108">To learn more about Excel features related to relationships between cells, see [Display the relationships between formulas and cells](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507).</span></span>

<span data-ttu-id="0399f-109">Une cellule peut avoir une cellule précédente et cette cellule peut avoir ses propres cellules précédentes.</span><span class="sxs-lookup"><span data-stu-id="0399f-109">A cell may have a precedent cell, and that precedent cell may have its own precedent cells.</span></span> <span data-ttu-id="0399f-110">Un « précédent direct » est le premier groupe de cellules précédent dans cette séquence, similaire au concept de parents dans une relation parent-enfant.</span><span class="sxs-lookup"><span data-stu-id="0399f-110">A "direct precedent" is the first preceding group of cells in this sequence, similar to the concept of parents in a parent-child relationship.</span></span> <span data-ttu-id="0399f-111">Un « dépendant direct » est le premier groupe dépendant de cellules dans une séquence, semblable aux enfants d’une relation parent-enfant.</span><span class="sxs-lookup"><span data-stu-id="0399f-111">A "direct dependent" is the first dependent group of cells in a sequence, similar to children in a parent-child relationship.</span></span> <span data-ttu-id="0399f-112">Les cellules qui font référence à d’autres cellules d’un workbook, mais dont la relation n’est pas une relation parent-enfant, ne sont pas des dépendants directs ou des antécédents directs.</span><span class="sxs-lookup"><span data-stu-id="0399f-112">Cells that refer to other cells in a workbook, but whose relationship is not a parent-child relationship, are not direct dependents or direct precedents.</span></span>

<span data-ttu-id="0399f-113">Cet article fournit des exemples de code qui récupèrent des antécédents directs et des dépendances directes des formules à l’aide de l Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="0399f-113">This article provides code samples that retrieve direct precedents and direct dependents of formulas using the Excel JavaScript API.</span></span> <span data-ttu-id="0399f-114">Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` [Range Object (interface API JavaScript pour Excel).](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="0399f-114">For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="get-the-direct-precedents-of-a-formula"></a><span data-ttu-id="0399f-115">Obtenir les antécédents directs d’une formule</span><span class="sxs-lookup"><span data-stu-id="0399f-115">Get the direct precedents of a formula</span></span>

<span data-ttu-id="0399f-116">Recherchez les cellules précédentes directes d’une formule [avec Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--).</span><span class="sxs-lookup"><span data-stu-id="0399f-116">Locate a formula's direct precedent cells with [Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--).</span></span> <span data-ttu-id="0399f-117">`Range.getDirectPrecedents` renvoie un `WorkbookRangeAreas` objet.</span><span class="sxs-lookup"><span data-stu-id="0399f-117">`Range.getDirectPrecedents` returns a `WorkbookRangeAreas` object.</span></span> <span data-ttu-id="0399f-118">Cet objet contient les adresses de tous les précédents directs du manuel.</span><span class="sxs-lookup"><span data-stu-id="0399f-118">This object contains the addresses of all the direct precedents in the workbook.</span></span> <span data-ttu-id="0399f-119">Il possède un objet `RangeAreas` distinct pour chaque feuille de calcul contenant au moins un précédent de formule.</span><span class="sxs-lookup"><span data-stu-id="0399f-119">It has a separate `RangeAreas` object for each worksheet containing at least one formula precedent.</span></span> <span data-ttu-id="0399f-120">Pour plus d’informations sur l’utilisation de l’objet, voir Work `RangeAreas` [with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="0399f-120">For more information on working with the `RangeAreas` object, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

<span data-ttu-id="0399f-121">La capture d’écran suivante montre le résultat de la sélection du bouton **Suivi des antécédents** dans Excel’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="0399f-121">The following screenshot shows the result of selecting the **Trace Precedents** button in the Excel UI.</span></span> <span data-ttu-id="0399f-122">Ce bouton dessine une flèche entre les cellules précédentes et la cellule sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="0399f-122">This button draws an arrow from precedent cells to the selected cell.</span></span> <span data-ttu-id="0399f-123">La cellule sélectionnée, **E3,** contient la formule « =C3 \* D3 », c’est pourquoi **C3** et **D3** sont des cellules précédentes.</span><span class="sxs-lookup"><span data-stu-id="0399f-123">The selected cell, **E3**, contains the formula "=C3 \* D3", so both **C3** and **D3** are precedent cells.</span></span> <span data-ttu-id="0399f-124">Contrairement au bouton Excel’interface utilisateur, `getDirectPrecedents` la méthode ne dessine pas de flèches.</span><span class="sxs-lookup"><span data-stu-id="0399f-124">Unlike the Excel UI button, the `getDirectPrecedents` method does not draw arrows.</span></span>

![Cellules précédentes de suivi des flèches dans Excel’interface utilisateur.](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> <span data-ttu-id="0399f-126">La `getDirectPrecedents` méthode ne peut pas récupérer les cellules précédentes dans les workbooks.</span><span class="sxs-lookup"><span data-stu-id="0399f-126">The `getDirectPrecedents` method can't retrieve precedent cells across workbooks.</span></span>

<span data-ttu-id="0399f-127">L’exemple de code suivant obtient les antécédents directs de la plage active, puis modifie la couleur d’arrière-plan de ces cellules précédentes en jaune.</span><span class="sxs-lookup"><span data-stu-id="0399f-127">The following code sample gets the direct precedents for the active range and then changes the background color of those precedent cells to yellow.</span></span>

```js
Excel.run(function (context) {
    // Precedents are cells that provide data to the selected formula.
    var range = context.workbook.getActiveCell();
    var directPrecedents = range.getDirectPrecedents();
    range.load("address");
    directPrecedents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct precedent cells of ${range.address}:`);

            // Use the direct precedents API to loop through precedents of the active cell.
            for (var i = 0; i < directPrecedents.areas.items.length; i++) {
              // Highlight and print out the address of each precedent cell.
              directPrecedents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directPrecedents.areas.items[i].address}`);
            }
        });
}).catch(errorHandlerFunction);
```

## <a name="get-the-direct-dependents-of-a-formula"></a><span data-ttu-id="0399f-128">Obtenir les dépendants directs d’une formule</span><span class="sxs-lookup"><span data-stu-id="0399f-128">Get the direct dependents of a formula</span></span>

<span data-ttu-id="0399f-129">Recherchez les cellules dépendantes directes d’une formule [avec Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__).</span><span class="sxs-lookup"><span data-stu-id="0399f-129">Locate a formula's direct dependent cells with [Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__).</span></span> <span data-ttu-id="0399f-130">Like `Range.getDirectPrecedents` , renvoie également un `Range.getDirectDependents` `WorkbookRangeAreas` objet.</span><span class="sxs-lookup"><span data-stu-id="0399f-130">Like `Range.getDirectPrecedents`, `Range.getDirectDependents` also returns a `WorkbookRangeAreas` object.</span></span> <span data-ttu-id="0399f-131">Cet objet contient les adresses de tous les dépendants directs dans le manuel.</span><span class="sxs-lookup"><span data-stu-id="0399f-131">This object contains the addresses of all the direct dependents in the workbook.</span></span> <span data-ttu-id="0399f-132">Il possède un objet `RangeAreas` distinct pour chaque feuille de calcul contenant au moins une formule dépendante.</span><span class="sxs-lookup"><span data-stu-id="0399f-132">It has a separate `RangeAreas` object for each worksheet containing at least one formula dependent.</span></span> <span data-ttu-id="0399f-133">Pour plus d’informations sur l’utilisation de l’objet, voir Work `RangeAreas` [with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="0399f-133">For more information on working with the `RangeAreas` object, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

<span data-ttu-id="0399f-134">La capture d’écran suivante montre le résultat de la sélection du bouton **Dépendants** du suivi dans Excel’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="0399f-134">The following screenshot shows the result of selecting the **Trace Dependents** button in the Excel UI.</span></span> <span data-ttu-id="0399f-135">Ce bouton dessine une flèche entre les cellules dépendantes et la cellule sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="0399f-135">This button draws an arrow from dependent cells to the selected cell.</span></span> <span data-ttu-id="0399f-136">La cellule sélectionnée, **D3,** a la cellule **E3** comme dépendant.</span><span class="sxs-lookup"><span data-stu-id="0399f-136">The selected cell, **D3**, has cell **E3** as a dependent.</span></span> <span data-ttu-id="0399f-137">**E3** contient la formule « =C3 \* D3 ».</span><span class="sxs-lookup"><span data-stu-id="0399f-137">**E3** contains the formula "=C3 \* D3".</span></span> <span data-ttu-id="0399f-138">Contrairement au bouton Excel’interface utilisateur, `getDirectDependents` la méthode ne dessine pas de flèches.</span><span class="sxs-lookup"><span data-stu-id="0399f-138">Unlike the Excel UI button, the `getDirectDependents` method does not draw arrows.</span></span>

![Cellules dépendantes du suivi des flèches dans Excel’interface utilisateur.](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> <span data-ttu-id="0399f-140">La `getDirectDependents` méthode ne peut pas récupérer les cellules dépendantes dans les workbooks.</span><span class="sxs-lookup"><span data-stu-id="0399f-140">The `getDirectDependents` method can't retrieve dependent cells across workbooks.</span></span>

<span data-ttu-id="0399f-141">L’exemple de code suivant obtient les dépendants directs de la plage active, puis modifie la couleur d’arrière-plan de ces cellules dépendantes en jaune.</span><span class="sxs-lookup"><span data-stu-id="0399f-141">The following code sample gets the direct dependents for the active range and then changes the background color of those dependent cells to yellow.</span></span>

```js
Excel.run(function (context) {
    // Direct dependents are cells that contain formulas that refer to other cells.
    var range = context.workbook.getActiveCell();
    var directDependents = range.getDirectDependents();
    range.load("address");
    directDependents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct dependent cells of ${range.address}:`);
    
            // Use the direct dependents API to loop through direct dependents of the active cell.
            for (var i = 0; i < directDependents.areas.items.length; i++) {
              // Highlight and print the address of each dependent cell.
              directDependents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directDependents.areas.items[i].address}`);
            }
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a><span data-ttu-id="0399f-142">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="0399f-142">See also</span></span>

- [<span data-ttu-id="0399f-143">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="0399f-143">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="0399f-144">Utiliser des cellules à l’aide Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="0399f-144">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="0399f-145">Travailler simultanément avec plusieurs plages dans des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="0399f-145">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
