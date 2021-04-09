---
title: Utiliser des antécédents de formule à l’aide de l’API JavaScript pour Excel
description: Découvrez comment utiliser l’API JavaScript excel pour récupérer les antécédents de formule.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0d21ae411615a22873a0f4dda185984f6191ac8e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652840"
---
# <a name="get-formula-precedents-using-the-excel-javascript-api"></a><span data-ttu-id="78aa0-103">Obtenir des antécédents de formule à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="78aa0-103">Get formula precedents using the Excel JavaScript API</span></span>

<span data-ttu-id="78aa0-104">Cet article fournit un exemple de code qui récupère les antécédents de formule à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="78aa0-104">This article provides a code sample that retrieves formula precedents using the Excel JavaScript API.</span></span> <span data-ttu-id="78aa0-105">Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` la classe [Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="78aa0-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="get-formula-precedents"></a><span data-ttu-id="78aa0-106">Obtenir des antécédents de formule</span><span class="sxs-lookup"><span data-stu-id="78aa0-106">Get formula precedents</span></span>

<span data-ttu-id="78aa0-107">Une formule Excel fait souvent référence à d’autres cellules.</span><span class="sxs-lookup"><span data-stu-id="78aa0-107">An Excel formula often refers to other cells.</span></span> <span data-ttu-id="78aa0-108">Lorsqu’une cellule fournit des données à une formule, elle est appelée formule « antécédent ».</span><span class="sxs-lookup"><span data-stu-id="78aa0-108">When a cell provides data to a formula, it is known as a formula "precedent".</span></span> <span data-ttu-id="78aa0-109">Pour en savoir plus sur les fonctionnalités Excel liées aux relations entre les cellules, voir Afficher les [relations entre les formules et les cellules.](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507)</span><span class="sxs-lookup"><span data-stu-id="78aa0-109">To learn more about Excel features related to relationships between cells, see [Display the relationships between formulas and cells](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507).</span></span> 

<span data-ttu-id="78aa0-110">Avec [Range.getDirectPrecedents,](/javascript/api/excel/excel.range#getdirectprecedents--)votre add-in peut localiser les cellules précédentes directes d’une formule.</span><span class="sxs-lookup"><span data-stu-id="78aa0-110">With [Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--), your add-in can locate a formula's direct precedent cells.</span></span> <span data-ttu-id="78aa0-111">`Range.getDirectPrecedents` renvoie un `WorkbookRangeAreas` objet.</span><span class="sxs-lookup"><span data-stu-id="78aa0-111">`Range.getDirectPrecedents` returns a `WorkbookRangeAreas` object.</span></span> <span data-ttu-id="78aa0-112">Cet objet contient les adresses de tous les antécédents dans le manuel.</span><span class="sxs-lookup"><span data-stu-id="78aa0-112">This object contains the addresses of all the precedents in the workbook.</span></span> <span data-ttu-id="78aa0-113">Il possède un objet `RangeAreas` distinct pour chaque feuille de calcul contenant au moins un précédent de formule.</span><span class="sxs-lookup"><span data-stu-id="78aa0-113">It has a separate `RangeAreas` object for each worksheet containing at least one formula precedent.</span></span> <span data-ttu-id="78aa0-114">Pour plus d’informations sur l’utilisation de l’objet, voir Work `RangeAreas` [with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="78aa0-114">For more information on working with the `RangeAreas` object, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

<span data-ttu-id="78aa0-115">Dans l’interface utilisateur Excel, le bouton **Suivi des antécédents** dessine une flèche entre les cellules précédentes et la formule sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="78aa0-115">In the Excel UI, the **Trace Precedents** button draws an arrow from precedent cells to the selected formula.</span></span> <span data-ttu-id="78aa0-116">Contrairement au bouton de l’interface utilisateur Excel, `getDirectPrecedents` la méthode ne dessine pas de flèches.</span><span class="sxs-lookup"><span data-stu-id="78aa0-116">Unlike the Excel UI button, the `getDirectPrecedents` method does not draw arrows.</span></span> 

> [!IMPORTANT]
> <span data-ttu-id="78aa0-117">La `getDirectPrecedents` méthode ne peut pas récupérer les cellules précédentes dans les workbooks.</span><span class="sxs-lookup"><span data-stu-id="78aa0-117">The `getDirectPrecedents` method can't retrieve precedent cells across workbooks.</span></span> 

<span data-ttu-id="78aa0-118">L’exemple de code suivant obtient les antécédents directs de la plage active, puis modifie la couleur d’arrière-plan de ces cellules précédentes en jaune.</span><span class="sxs-lookup"><span data-stu-id="78aa0-118">The following code sample gets the direct precedents for the active range and then changes the background color of those precedent cells to yellow.</span></span> 

> [!NOTE]
> <span data-ttu-id="78aa0-119">La plage active doit contenir une formule qui fait référence à d’autres cellules du même workbook pour que la mise en surbrillance fonctionne correctement.</span><span class="sxs-lookup"><span data-stu-id="78aa0-119">The active range must contain a formula that references other cells in the same workbook for the highlighting to work properly.</span></span> 

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
        })
        .then(context.sync);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a><span data-ttu-id="78aa0-120">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="78aa0-120">See also</span></span>

- [<span data-ttu-id="78aa0-121">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="78aa0-121">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="78aa0-122">Utiliser des cellules à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="78aa0-122">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="78aa0-123">Travailler simultanément avec plusieurs plages dans des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="78aa0-123">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
