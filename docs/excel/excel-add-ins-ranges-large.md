---
title: Lire ou écrire dans de grandes plages à l’aide de l’API JavaScript pour Excel
description: Découvrez comment lire ou écrire dans de grandes plages avec l’API JavaScript pour Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: b7a1e54d6b516889884f777bd256df8fb663c794
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652841"
---
# <a name="read-or-write-to-a-large-range-using-the-excel-javascript-api"></a><span data-ttu-id="1acad-103">Lire ou écrire dans une grande plage à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="1acad-103">Read or write to a large range using the Excel JavaScript API</span></span>

<span data-ttu-id="1acad-104">Cet article explique comment gérer la lecture et l’écriture dans de grandes plages avec l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="1acad-104">This article describes how to handle reading and writing to large ranges with the Excel JavaScript API.</span></span>

## <a name="run-separate-read-or-write-operations-for-large-ranges"></a><span data-ttu-id="1acad-105">Exécuter des opérations de lecture ou d’écriture distinctes pour des plages de grande taille</span><span class="sxs-lookup"><span data-stu-id="1acad-105">Run separate read or write operations for large ranges</span></span>

<span data-ttu-id="1acad-106">Si une plage contient un grand nombre de cellules, valeurs, formats numériques ou formules, il est possible qu’il ne soit pas possible d’exécuter des opérations API sur cette plage.</span><span class="sxs-lookup"><span data-stu-id="1acad-106">If a range contains a large number of cells, values, number formats, or formulas, it may not be possible to run API operations on that range.</span></span> <span data-ttu-id="1acad-107">L’API essaie toujours d’exécuter au mieux l’opération demandée sur une plage (par exemple, pour extraire ou écrire des données spécifiées), mais essayer d’effectuer des opérations de lecture ou d’écriture pour une grande plage peut provoquer une erreur d’API en raison de l’utilisation des ressources excessive.</span><span class="sxs-lookup"><span data-stu-id="1acad-107">The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization.</span></span> <span data-ttu-id="1acad-108">Pour éviter ces erreurs, nous vous recommandons d’exécuter des opérations de lecture ou d’écriture distinctes pour des sous-ensembles plus petits d’une grande plage, au lieu d’essayer d’exécuter une seule opération de lecture ou d’écriture sur une grande plage.</span><span class="sxs-lookup"><span data-stu-id="1acad-108">To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>

<span data-ttu-id="1acad-109">Pour plus d’informations sur les limitations du système, voir la section « Excel add-ins » des limites de ressources et l’optimisation des performances pour les [add-ins Office.](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)</span><span class="sxs-lookup"><span data-stu-id="1acad-109">For details on the system limitations, see the "Excel add-ins" section of [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).</span></span>

### <a name="conditional-formatting-of-ranges"></a><span data-ttu-id="1acad-110">Mise en forme conditionnelle de plages</span><span class="sxs-lookup"><span data-stu-id="1acad-110">Conditional formatting of ranges</span></span>

<span data-ttu-id="1acad-111">Des plages peuvent présenter une mise en forme de cellules individuelles en fonction de certaines conditions.</span><span class="sxs-lookup"><span data-stu-id="1acad-111">Ranges can have formats applied to individual cells based on conditions.</span></span> <span data-ttu-id="1acad-112">Pour plus d’informations à ce sujet, consultez l’article [Appliquer une mise en forme conditionnelle à des plages Excel](excel-add-ins-conditional-formatting.md).</span><span class="sxs-lookup"><span data-stu-id="1acad-112">For more information about this, see [Apply conditional formatting to Excel ranges](excel-add-ins-conditional-formatting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="1acad-113">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="1acad-113">See also</span></span>

- [<span data-ttu-id="1acad-114">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="1acad-114">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="1acad-115">Utiliser des cellules à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="1acad-115">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="1acad-116">Lire ou écrire dans une plage non limite à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="1acad-116">Read or write to an unbounded range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-unbounded.md)
- [<span data-ttu-id="1acad-117">Travailler simultanément avec plusieurs plages dans des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="1acad-117">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
