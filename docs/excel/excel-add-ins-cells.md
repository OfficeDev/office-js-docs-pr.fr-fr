---
title: Utiliser des cellules à l'aide de l'API JavaScript pour Excel.
description: Découvrez la définition de l'API JavaScript pour Excel d'une cellule et découvrez comment utiliser des cellules.
ms.date: 04/16/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad8ca985b6bbdcf19920c36c371e690f61639f16
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917099"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a><span data-ttu-id="c2b55-103">Utiliser des cellules à l'aide de l'API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="c2b55-103">Work with cells using the Excel JavaScript API</span></span>

<span data-ttu-id="c2b55-104">L’API JavaScript Excel ne comprend pas d’objet ou de classe « Cellule ».</span><span class="sxs-lookup"><span data-stu-id="c2b55-104">The Excel JavaScript API doesn't have a "Cell" object or class.</span></span> <span data-ttu-id="c2b55-105">Au lieu de cela, toutes les cellules Excel sont `Range` des objets.</span><span class="sxs-lookup"><span data-stu-id="c2b55-105">Instead, all Excel cells are `Range` objects.</span></span> <span data-ttu-id="c2b55-106">Une cellule individuelle dans l’interface utilisateur d’Excel se traduit par un objet`Range` avec une cellule dans l’API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="c2b55-106">An individual cell in the Excel UI translates to a `Range` object with one cell in the Excel JavaScript API.</span></span>

<span data-ttu-id="c2b55-107">Un `Range` objet peut également contenir plusieurs cellules contiguës.</span><span class="sxs-lookup"><span data-stu-id="c2b55-107">A `Range` object can also contain multiple, contiguous cells.</span></span> <span data-ttu-id="c2b55-108">Les cellules contiguës forment un rectangle non abandonné (y compris des lignes ou des colonnes).</span><span class="sxs-lookup"><span data-stu-id="c2b55-108">Contiguous cells form an unbroken rectangle (including single rows or columns).</span></span> <span data-ttu-id="c2b55-109">Pour en savoir plus sur l'utilisation de cellules qui ne sont pas contiguës, voir Travailler avec des cellules non contiguës à l'aide de l'objet [RangeAreas](#work-with-discontiguous-cells-using-the-rangeareas-object).</span><span class="sxs-lookup"><span data-stu-id="c2b55-109">To learn about working with cells that are not contiguous, see [Work with discontiguous cells using the RangeAreas object](#work-with-discontiguous-cells-using-the-rangeareas-object).</span></span>

<span data-ttu-id="c2b55-110">Pour obtenir la liste complète des propriétés et méthodes que l'objet prend en charge, voir `Range` [Range Object (interface API JavaScript pour Excel).](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="c2b55-110">For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a><span data-ttu-id="c2b55-111">Utiliser des cellules peuigues à l'aide de l'objet RangeAreas</span><span class="sxs-lookup"><span data-stu-id="c2b55-111">Work with discontiguous cells using the RangeAreas object</span></span>

<span data-ttu-id="c2b55-112">[L'objet RangeAreas permet](/javascript/api/excel/excel.rangeareas) à votre add-in d'effectuer des opérations sur plusieurs plages à la fois.</span><span class="sxs-lookup"><span data-stu-id="c2b55-112">The [RangeAreas](/javascript/api/excel/excel.rangeareas) object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="c2b55-113">Ces plages peuvent être contiguës, mais elles n'en ont pas besoin.</span><span class="sxs-lookup"><span data-stu-id="c2b55-113">These ranges may be contiguous, but they don't have to be.</span></span> <span data-ttu-id="c2b55-114">`RangeAreas`sont abordés plus loin dans l’article[Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="c2b55-114">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="c2b55-115">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c2b55-115">See also</span></span>

- [<span data-ttu-id="c2b55-116">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="c2b55-116">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="c2b55-117">Obtenir une plage à l'aide de l'API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="c2b55-117">Get a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-get.md)
- [<span data-ttu-id="c2b55-118">Travailler simultanément avec plusieurs plages dans des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="c2b55-118">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
