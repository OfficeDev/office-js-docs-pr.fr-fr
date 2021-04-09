---
title: Utiliser des cellules à l’aide de l’API JavaScript pour Excel.
description: Découvrez la définition de l’API JavaScript pour Excel d’une cellule et découvrez comment utiliser des cellules.
ms.date: 04/07/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 5fcfeeef52f17c22d13ed3c1a10851f1d8e69204
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652873"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a><span data-ttu-id="98979-103">Utiliser des cellules à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="98979-103">Work with cells using the Excel JavaScript API</span></span>

<span data-ttu-id="98979-104">L’API JavaScript pour Excel n’a pas d’objet ou de classe « Cell ».</span><span class="sxs-lookup"><span data-stu-id="98979-104">The Excel JavaScript API doesn't have a "Cell" object or class.</span></span> <span data-ttu-id="98979-105">Au lieu de cela, toutes les cellules Excel sont `Range` des objets.</span><span class="sxs-lookup"><span data-stu-id="98979-105">Instead, all Excel cells are `Range` objects.</span></span> <span data-ttu-id="98979-106">Une cellule individuelle dans l’interface utilisateur d’Excel se traduit par un objet avec une cellule dans `Range` l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="98979-106">An individual cell in the Excel UI translates to a `Range` object with one cell in the Excel JavaScript API.</span></span>

<span data-ttu-id="98979-107">Un `Range` objet peut également contenir plusieurs cellules contiguës.</span><span class="sxs-lookup"><span data-stu-id="98979-107">A `Range` object can also contain multiple, contiguous cells.</span></span> <span data-ttu-id="98979-108">Les cellules contiguës forment un rectangle non abandonné (y compris des lignes ou des colonnes).</span><span class="sxs-lookup"><span data-stu-id="98979-108">Contiguous cells form an unbroken rectangle (including single rows or columns).</span></span> <span data-ttu-id="98979-109">Pour en savoir plus sur l’utilisation de cellules qui ne sont pas contiguës, voir Travailler avec des cellules non contiguës à l’aide de l’objet [RangeAreas](#work-with-discontiguous-cells-using-the-rangeareas-object).</span><span class="sxs-lookup"><span data-stu-id="98979-109">To learn about working with cells that are not contiguous, see [Work with discontiguous cells using the RangeAreas object](#work-with-discontiguous-cells-using-the-rangeareas-object).</span></span>

<span data-ttu-id="98979-110">Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` la classe [Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="98979-110">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="excel-javascript-apis-that-mention-cells"></a><span data-ttu-id="98979-111">API JavaScript Excel mentionnant les cellules</span><span class="sxs-lookup"><span data-stu-id="98979-111">Excel JavaScript APIs that mention cells</span></span>

<span data-ttu-id="98979-112">Même si l’API JavaScript pour Excel n’a pas d’objet ou de classe « Cell », un certain nombre de noms d’API mentionnent des cellules.</span><span class="sxs-lookup"><span data-stu-id="98979-112">Even though the Excel JavaScript API doesn't have a "Cell" object or class, a number of API names mention cells.</span></span> <span data-ttu-id="98979-113">Ces API contrôlent les propriétés des cellules telles que la couleur, la mise en forme du texte et la police.</span><span class="sxs-lookup"><span data-stu-id="98979-113">These APIs control cell properties like color, text formatting, and font.</span></span>

<span data-ttu-id="98979-114">La liste suivante des API JavaScript pour Excel fait référence à des cellules.</span><span class="sxs-lookup"><span data-stu-id="98979-114">The following list of Excel JavaScript APIs refer to cells.</span></span>

- [<span data-ttu-id="98979-115">CellBorder</span><span class="sxs-lookup"><span data-stu-id="98979-115">CellBorder</span></span>](/javascript/api/excel/excel.cellborder)
- [<span data-ttu-id="98979-116">CellBorderCollection</span><span class="sxs-lookup"><span data-stu-id="98979-116">CellBorderCollection</span></span>](/javascript/api/excel/excel.cellbordercollection)
- [<span data-ttu-id="98979-117">CellProperties</span><span class="sxs-lookup"><span data-stu-id="98979-117">CellProperties</span></span>](/javascript/api/excel/excel.cellproperties)
- [<span data-ttu-id="98979-118">CellPropertiesFill</span><span class="sxs-lookup"><span data-stu-id="98979-118">CellPropertiesFill</span></span>](/javascript/api/excel/excel.cellpropertiesfill)
- [<span data-ttu-id="98979-119">CellPropertiesFont</span><span class="sxs-lookup"><span data-stu-id="98979-119">CellPropertiesFont</span></span>](/javascript/api/excel/excel.cellpropertiesfont)
- [<span data-ttu-id="98979-120">CellPropertiesFormat</span><span class="sxs-lookup"><span data-stu-id="98979-120">CellPropertiesFormat</span></span>](/javascript/api/excel/excel.cellpropertiesformat)
- [<span data-ttu-id="98979-121">CellPropertiesProtection</span><span class="sxs-lookup"><span data-stu-id="98979-121">CellPropertiesProtection</span></span>](/javascript/api/excel/excel.cellpropertiesprotection)
- [<span data-ttu-id="98979-122">CellValueConditionalFormat</span><span class="sxs-lookup"><span data-stu-id="98979-122">CellValueConditionalFormat</span></span>](/javascript/api/excel/excel.cellvalueconditionalformat)
- [<span data-ttu-id="98979-123">ConditionalCellValueRule</span><span class="sxs-lookup"><span data-stu-id="98979-123">ConditionalCellValueRule</span></span>](/javascript/api/excel/excel.conditionalcellvaluerule)
- [<span data-ttu-id="98979-124">SettableCellProperties</span><span class="sxs-lookup"><span data-stu-id="98979-124">SettableCellProperties</span></span>](/javascript/api/excel/excel.settablecellproperties)

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a><span data-ttu-id="98979-125">Utiliser des cellules peuigues à l’aide de l’objet RangeAreas</span><span class="sxs-lookup"><span data-stu-id="98979-125">Work with discontiguous cells using the RangeAreas object</span></span>

<span data-ttu-id="98979-126">[L’objet RangeAreas permet](/javascript/api/excel/excel.rangeareas) à votre add-in d’effectuer des opérations sur plusieurs plages à la fois.</span><span class="sxs-lookup"><span data-stu-id="98979-126">The [RangeAreas](/javascript/api/excel/excel.rangeareas) object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="98979-127">Ces plages peuvent être contiguës, mais elles n’en ont pas besoin.</span><span class="sxs-lookup"><span data-stu-id="98979-127">These ranges may be contiguous, but they don't have to be.</span></span> <span data-ttu-id="98979-128">`RangeAreas`sont abordés plus loin dans l’article[Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="98979-128">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="98979-129">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="98979-129">See also</span></span>

- [<span data-ttu-id="98979-130">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="98979-130">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="98979-131">Obtenir une plage à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="98979-131">Get a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-get.md)
- [<span data-ttu-id="98979-132">Travailler simultanément avec plusieurs plages dans des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="98979-132">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
