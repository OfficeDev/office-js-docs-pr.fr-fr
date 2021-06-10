---
title: Utiliser des antécédents et des dépendances de formule à l’aide Excel API JavaScript
description: Découvrez comment utiliser l’API JavaScript Excel pour récupérer les antécédents et les dépendances de formule.
ms.date: 06/03/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 78fa4fb070ede85d139425a9d59ba1224785a605
ms.sourcegitcommit: 17b5a076375bc5dc3f91d3602daeb7535d67745d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/06/2021
ms.locfileid: "52783522"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a><span data-ttu-id="108e9-103">Obtenir des antécédents et des dépendances de formule à l’aide Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="108e9-103">Get formula precedents and dependents using the Excel JavaScript API</span></span>

<span data-ttu-id="108e9-104">Excel formules font souvent référence à d’autres cellules.</span><span class="sxs-lookup"><span data-stu-id="108e9-104">Excel formulas often refer to other cells.</span></span> <span data-ttu-id="108e9-105">Ces références entre cellules sont appelées « antécédents » et « dépendants ».</span><span class="sxs-lookup"><span data-stu-id="108e9-105">These cross-cell references are known as "precedents" and "dependents".</span></span> <span data-ttu-id="108e9-106">Un précédent est une cellule qui fournit des données à une formule.</span><span class="sxs-lookup"><span data-stu-id="108e9-106">A precedent is a cell that provides data to a formula.</span></span> <span data-ttu-id="108e9-107">Une cellule dépendante est une cellule qui contient une formule qui fait référence à d’autres cellules.</span><span class="sxs-lookup"><span data-stu-id="108e9-107">A dependent is a cell that contains a formula that refers to other cells.</span></span> <span data-ttu-id="108e9-108">Pour en savoir plus sur Excel fonctionnalités liées aux relations entre les cellules, voir Afficher les relations entre les [formules et les cellules.](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507)</span><span class="sxs-lookup"><span data-stu-id="108e9-108">To learn more about Excel features related to relationships between cells, see [Display the relationships between formulas and cells](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507).</span></span>

<span data-ttu-id="108e9-109">Une cellule peut avoir une cellule précédente et cette cellule peut avoir ses propres cellules précédentes.</span><span class="sxs-lookup"><span data-stu-id="108e9-109">A cell may have a precedent cell, and that precedent cell may have its own precedent cells.</span></span> <span data-ttu-id="108e9-110">Un « précédent direct » est le premier groupe de cellules précédent dans cette séquence, similaire au concept de parents dans une relation parent-enfant.</span><span class="sxs-lookup"><span data-stu-id="108e9-110">A "direct precedent" is the first preceding group of cells in this sequence, similar to the concept of parents in a parent-child relationship.</span></span> <span data-ttu-id="108e9-111">Un « dépendant direct » est le premier groupe dépendant de cellules dans une séquence, semblable aux enfants d’une relation parent-enfant.</span><span class="sxs-lookup"><span data-stu-id="108e9-111">A "direct dependent" is the first dependent group of cells in a sequence, similar to children in a parent-child relationship.</span></span> <span data-ttu-id="108e9-112">Les cellules qui font référence à d’autres cellules d’un workbook, mais dont la relation n’est pas une relation parent-enfant, ne sont pas des dépendants directs ou des antécédents directs.</span><span class="sxs-lookup"><span data-stu-id="108e9-112">Cells that refer to other cells in a workbook, but whose relationship is not a parent-child relationship, are not direct dependents or direct precedents.</span></span>

<span data-ttu-id="108e9-113">Cet article fournit des exemples de code qui récupèrent les antécédents directs et les dépendances directes des formules à l’aide Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="108e9-113">This article provides code samples that retrieve direct precedents and direct dependents of formulas using the Excel JavaScript API.</span></span> <span data-ttu-id="108e9-114">Pour obtenir la liste complète des propriétés et méthodes que l’objet prend en charge, voir `Range` [Range Object (interface API JavaScript pour Excel).](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="108e9-114">For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="get-the-direct-precedents-of-a-formula"></a><span data-ttu-id="108e9-115">Obtenir les antécédents directs d’une formule</span><span class="sxs-lookup"><span data-stu-id="108e9-115">Get the direct precedents of a formula</span></span>

<span data-ttu-id="108e9-116">Recherchez les cellules précédentes directes d’une formule [avec Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--).</span><span class="sxs-lookup"><span data-stu-id="108e9-116">Locate a formula's direct precedent cells with [Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--).</span></span> <span data-ttu-id="108e9-117">`Range.getDirectPrecedents` renvoie un `WorkbookRangeAreas` objet.</span><span class="sxs-lookup"><span data-stu-id="108e9-117">`Range.getDirectPrecedents` returns a `WorkbookRangeAreas` object.</span></span> <span data-ttu-id="108e9-118">Cet objet contient les adresses de tous les précédents directs du manuel.</span><span class="sxs-lookup"><span data-stu-id="108e9-118">This object contains the addresses of all the direct precedents in the workbook.</span></span> <span data-ttu-id="108e9-119">Il possède un objet `RangeAreas` distinct pour chaque feuille de calcul contenant au moins un précédent de formule.</span><span class="sxs-lookup"><span data-stu-id="108e9-119">It has a separate `RangeAreas` object for each worksheet containing at least one formula precedent.</span></span> <span data-ttu-id="108e9-120">Pour plus d’informations sur l’utilisation de l’objet, voir Work `RangeAreas` [with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="108e9-120">For more information on working with the `RangeAreas` object, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

<span data-ttu-id="108e9-121">La capture d’écran suivante montre le résultat de la sélection du bouton Suivi **des antécédents** dans Excel’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="108e9-121">The following screenshot shows the result of selecting the **Trace Precedents** button in the Excel UI.</span></span> <span data-ttu-id="108e9-122">Ce bouton dessine une flèche entre les cellules précédentes et la cellule sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="108e9-122">This button draws an arrow from precedent cells to the selected cell.</span></span> <span data-ttu-id="108e9-123">La cellule sélectionnée, **E3,** contient la formule « =C3 \* D3 », c’est pourquoi **C3** et **D3** sont des cellules précédentes.</span><span class="sxs-lookup"><span data-stu-id="108e9-123">The selected cell, **E3**, contains the formula "=C3 \* D3", so both **C3** and **D3** are precedent cells.</span></span> <span data-ttu-id="108e9-124">Contrairement au bouton Excel’interface utilisateur, `getDirectPrecedents` la méthode ne dessine pas de flèches.</span><span class="sxs-lookup"><span data-stu-id="108e9-124">Unlike the Excel UI button, the `getDirectPrecedents` method does not draw arrows.</span></span>

![Cellules précédentes de suivi des flèches dans l Excel’interface utilisateur](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> <span data-ttu-id="108e9-126">La `getDirectPrecedents` méthode ne peut pas récupérer les cellules précédentes dans les workbooks.</span><span class="sxs-lookup"><span data-stu-id="108e9-126">The `getDirectPrecedents` method can't retrieve precedent cells across workbooks.</span></span>

<span data-ttu-id="108e9-127">L’exemple de code suivant obtient les antécédents directs de la plage active, puis change la couleur d’arrière-plan de ces cellules précédentes en jaune.</span><span class="sxs-lookup"><span data-stu-id="108e9-127">The following code sample gets the direct precedents for the active range and then changes the background color of those precedent cells to yellow.</span></span>

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

## <a name="get-the-direct-dependents-of-a-formula-preview"></a><span data-ttu-id="108e9-128">Obtenir les dépendants directs d’une formule (aperçu)</span><span class="sxs-lookup"><span data-stu-id="108e9-128">Get the direct dependents of a formula (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="108e9-129">La `Range.getDirectDependents` méthode est actuellement disponible uniquement en prévisualisation publique.</span><span class="sxs-lookup"><span data-stu-id="108e9-129">The `Range.getDirectDependents` method is currently only available in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

<span data-ttu-id="108e9-130">Recherchez les cellules dépendantes directes d’une formule [avec Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__).</span><span class="sxs-lookup"><span data-stu-id="108e9-130">Locate a formula's direct dependent cells with [Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__).</span></span> <span data-ttu-id="108e9-131">Like `Range.getDirectPrecedents` , renvoie également un `Range.getDirectDependents` `WorkbookRangeAreas` objet.</span><span class="sxs-lookup"><span data-stu-id="108e9-131">Like `Range.getDirectPrecedents`, `Range.getDirectDependents` also returns a `WorkbookRangeAreas` object.</span></span> <span data-ttu-id="108e9-132">Cet objet contient les adresses de tous les dépendants directs dans le manuel.</span><span class="sxs-lookup"><span data-stu-id="108e9-132">This object contains the addresses of all the direct dependents in the workbook.</span></span> <span data-ttu-id="108e9-133">Il possède un objet `RangeAreas` distinct pour chaque feuille de calcul contenant au moins une formule dépendante.</span><span class="sxs-lookup"><span data-stu-id="108e9-133">It has a separate `RangeAreas` object for each worksheet containing at least one formula dependent.</span></span> <span data-ttu-id="108e9-134">Pour plus d’informations sur l’utilisation de l’objet, voir Work `RangeAreas` [with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="108e9-134">For more information on working with the `RangeAreas` object, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

<span data-ttu-id="108e9-135">La capture d’écran suivante montre le résultat de la sélection du bouton **Dépendants** du suivi dans Excel’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="108e9-135">The following screenshot shows the result of selecting the **Trace Dependents** button in the Excel UI.</span></span> <span data-ttu-id="108e9-136">Ce bouton dessine une flèche entre les cellules dépendantes et la cellule sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="108e9-136">This button draws an arrow from dependent cells to the selected cell.</span></span> <span data-ttu-id="108e9-137">La cellule sélectionnée, **D3,** a la cellule **E3** comme dépendant.</span><span class="sxs-lookup"><span data-stu-id="108e9-137">The selected cell, **D3**, has cell **E3** as a dependent.</span></span> <span data-ttu-id="108e9-138">**E3** contient la formule « =C3 \* D3 ».</span><span class="sxs-lookup"><span data-stu-id="108e9-138">**E3** contains the formula "=C3 \* D3".</span></span> <span data-ttu-id="108e9-139">Contrairement au bouton Excel’interface utilisateur, `getDirectDependents` la méthode ne dessine pas de flèches.</span><span class="sxs-lookup"><span data-stu-id="108e9-139">Unlike the Excel UI button, the `getDirectDependents` method does not draw arrows.</span></span>

![Cellules dépendantes du suivi des flèches dans l Excel’interface utilisateur](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> <span data-ttu-id="108e9-141">La `getDirectDependents` méthode ne peut pas récupérer les cellules dépendantes dans les workbooks.</span><span class="sxs-lookup"><span data-stu-id="108e9-141">The `getDirectDependents` method can't retrieve dependent cells across workbooks.</span></span>

<span data-ttu-id="108e9-142">L’exemple de code suivant obtient les dépendants directs de la plage active, puis modifie la couleur d’arrière-plan de ces cellules dépendantes en jaune.</span><span class="sxs-lookup"><span data-stu-id="108e9-142">The following code sample gets the direct dependents for the active range and then changes the background color of those dependent cells to yellow.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="108e9-143">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="108e9-143">See also</span></span>

- [<span data-ttu-id="108e9-144">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="108e9-144">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="108e9-145">Utiliser des cellules à l’aide de Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="108e9-145">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="108e9-146">Travailler simultanément avec plusieurs plages dans des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="108e9-146">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)