---
title: Plages de groupes à l’aide Excel API JavaScript
description: Découvrez comment grouper des lignes ou des colonnes d’une plage pour créer un plan à l’aide Excel API JavaScript.
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 960a394a1467ec1fe55ff8dbf7b0a3f39fd355a5
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075718"
---
# <a name="group-ranges-for-an-outline-using-the-excel-javascript-api"></a><span data-ttu-id="28911-103">Plages de groupe pour un plan à l’aide de l Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="28911-103">Group ranges for an outline using the Excel JavaScript API</span></span>

<span data-ttu-id="28911-104">Cet article fournit un exemple de code qui montre comment grouper des plages pour un plan à l’aide Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="28911-104">This article provides a code sample that shows how to group ranges for an outline using the Excel JavaScript API.</span></span> <span data-ttu-id="28911-105">Pour obtenir la liste complète des propriétés et méthodes que l’objet prend en `Range` charge, [voir Excel. Classe Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="28911-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="group-rows-or-columns-of-a-range-for-an-outline"></a><span data-ttu-id="28911-106">Grouper des lignes ou des colonnes d’une plage pour un plan</span><span class="sxs-lookup"><span data-stu-id="28911-106">Group rows or columns of a range for an outline</span></span>

<span data-ttu-id="28911-107">Les lignes ou colonnes d’une plage peuvent être regroupées pour créer un [plan.](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF)</span><span class="sxs-lookup"><span data-stu-id="28911-107">Rows or columns of a range can be grouped together to create an [outline](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF).</span></span> <span data-ttu-id="28911-108">Ces groupes peuvent être réduire et développés pour masquer et afficher les cellules correspondantes.</span><span class="sxs-lookup"><span data-stu-id="28911-108">These groups can be collapsed and expanded to hide and show the corresponding cells.</span></span> <span data-ttu-id="28911-109">Cela facilite l’analyse rapide des données de première ligne.</span><span class="sxs-lookup"><span data-stu-id="28911-109">This makes quick analysis of top-line data easier.</span></span> <span data-ttu-id="28911-110">Utilisez [Range.group pour](/javascript/api/excel/excel.range#group-groupoption-) effectuer ces groupes de plan.</span><span class="sxs-lookup"><span data-stu-id="28911-110">Use [Range.group](/javascript/api/excel/excel.range#group-groupoption-) to make these outline groups.</span></span>

<span data-ttu-id="28911-111">Un plan peut avoir une hiérarchie, où des groupes plus petits sont imbrmbrés sous des groupes plus grands.</span><span class="sxs-lookup"><span data-stu-id="28911-111">An outline can have a hierarchy, where smaller groups are nested under larger groups.</span></span> <span data-ttu-id="28911-112">Cela permet d’afficher le plan à différents niveaux.</span><span class="sxs-lookup"><span data-stu-id="28911-112">This allows the outline to be viewed at different levels.</span></span> <span data-ttu-id="28911-113">La modification du niveau de plan visible peut être effectuée par programme via la [méthode Worksheet.showOutlineLevels.](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-)</span><span class="sxs-lookup"><span data-stu-id="28911-113">Changing the visible outline level can be done programmatically through the [Worksheet.showOutlineLevels](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-) method.</span></span> <span data-ttu-id="28911-114">Notez que Excel ne prend en charge que huit niveaux de groupes de plan.</span><span class="sxs-lookup"><span data-stu-id="28911-114">Note that Excel only supports eight levels of outline groups.</span></span>

<span data-ttu-id="28911-115">L’exemple de code suivant crée un plan avec deux niveaux de groupes pour les lignes et les colonnes.</span><span class="sxs-lookup"><span data-stu-id="28911-115">The following code sample creates an outline with two levels of groups for both the rows and columns.</span></span> <span data-ttu-id="28911-116">L’image suivante montre les regroupements de ce plan.</span><span class="sxs-lookup"><span data-stu-id="28911-116">The subsequent image shows the groupings of that outline.</span></span> <span data-ttu-id="28911-117">Dans l’exemple de code, les plages regroupées n’incluent pas la ligne ou la colonne du contrôle de plan (les « Totaux » pour cet exemple).</span><span class="sxs-lookup"><span data-stu-id="28911-117">In the code sample, the ranges being grouped do not include the row or column of the outline control (the "Totals" for this example).</span></span> <span data-ttu-id="28911-118">Un groupe définit ce qui sera réduire, et non la ligne ou la colonne avec le contrôle.</span><span class="sxs-lookup"><span data-stu-id="28911-118">A group defines what will be collapsed, not the row or column with the control.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Group the larger, main level. Note that the outline controls
    // will be on row 10, meaning 4-9 will collapse and expand.
    sheet.getRange("4:9").group(Excel.GroupOption.byRows);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on rows 6 and 9, meaning 4-5 and 7-8 will collapse and expand.
    sheet.getRange("4:5").group(Excel.GroupOption.byRows);
    sheet.getRange("7:8").group(Excel.GroupOption.byRows);

    // Group the larger, main level. Note that the outline controls
    // will be on column R, meaning C-Q will collapse and expand.
    sheet.getRange("C:Q").group(Excel.GroupOption.byColumns);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on columns G, L, and R, meaning C-F, H-K, and M-P will collapse and expand.
    sheet.getRange("C:F").group(Excel.GroupOption.byColumns);
    sheet.getRange("H:K").group(Excel.GroupOption.byColumns);
    sheet.getRange("M:P").group(Excel.GroupOption.byColumns);
    return context.sync();
}).catch(errorHandlerFunction);
```

![Plage avec un plan à deux niveaux à deux dimensions.](../images/excel-outline.png)

## <a name="remove-grouping-from-rows-or-columns-of-a-range"></a><span data-ttu-id="28911-120">Supprimer le regroupement des lignes ou des colonnes d’une plage</span><span class="sxs-lookup"><span data-stu-id="28911-120">Remove grouping from rows or columns of a range</span></span>

<span data-ttu-id="28911-121">Pour regrouper un groupe de lignes ou de colonnes, utilisez [la méthode Range.ungroup.](/javascript/api/excel/excel.range#ungroup-groupoption-)</span><span class="sxs-lookup"><span data-stu-id="28911-121">To ungroup a row or column group, use the [Range.ungroup](/javascript/api/excel/excel.range#ungroup-groupoption-) method.</span></span> <span data-ttu-id="28911-122">Cela supprime le niveau le plus à l’extérieur du plan.</span><span class="sxs-lookup"><span data-stu-id="28911-122">This removes the outermost level from the outline.</span></span> <span data-ttu-id="28911-123">Si plusieurs groupes du même type de ligne ou de colonne sont au même niveau dans la plage spécifiée, tous ces groupes sont désgroupés.</span><span class="sxs-lookup"><span data-stu-id="28911-123">If multiple groups of the same row or column type are at the same level within the specified range, all of those groups are ungrouped.</span></span>

## <a name="see-also"></a><span data-ttu-id="28911-124">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="28911-124">See also</span></span>

- [<span data-ttu-id="28911-125">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="28911-125">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="28911-126">Utiliser des cellules à l’aide Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="28911-126">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="28911-127">Travailler simultanément avec plusieurs plages dans des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="28911-127">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
