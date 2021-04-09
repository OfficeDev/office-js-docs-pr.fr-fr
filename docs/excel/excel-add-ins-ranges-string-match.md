---
title: Rechercher une chaîne à l’aide de l’API JavaScript pour Excel
description: Découvrez comment trouver une chaîne dans une plage à l’aide de l’API JavaScript pour Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9b649bb249cd24d7578bc4f8285e5d0a23d0e4cd
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652821"
---
# <a name="find-a-string-within-a-range-using-the-excel-javascript-api"></a><span data-ttu-id="c1897-103">Rechercher une chaîne dans une plage à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="c1897-103">Find a string within a range using the Excel JavaScript API</span></span>

<span data-ttu-id="c1897-104">Cet article fournit un exemple de code qui trouve une chaîne dans une plage à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="c1897-104">This article provides a code sample that finds a string within a range using the Excel JavaScript API.</span></span> <span data-ttu-id="c1897-105">Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` la classe [Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="c1897-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="match-a-string-within-a-range"></a><span data-ttu-id="c1897-106">Faire correspondre une chaîne dans une plage</span><span class="sxs-lookup"><span data-stu-id="c1897-106">Match a string within a range</span></span>

<span data-ttu-id="c1897-107">L’objet `Range` dispose d’une méthode`find` pour rechercher une chaîne spécifiée dans la plage.</span><span class="sxs-lookup"><span data-stu-id="c1897-107">The `Range` object has a `find` method to search for a specified string within the range.</span></span> <span data-ttu-id="c1897-108">Elle renvoie la plage de la première cellule avec le texte correspondant.</span><span class="sxs-lookup"><span data-stu-id="c1897-108">It returns the range of the first cell with matching text.</span></span>

<span data-ttu-id="c1897-109">L’exemple de code suivant trouve la première cellule contenant une valeur égale à la chaîne **Nourriture** et connecte son adresse à la console.</span><span class="sxs-lookup"><span data-stu-id="c1897-109">The following code sample finds the first cell with a value equal to the string **Food** and logs its address to the console.</span></span> <span data-ttu-id="c1897-110">Notez que `find` génère une erreur `ItemNotFound` si la chaîne spécifiée n’existe pas dans la plage.</span><span class="sxs-lookup"><span data-stu-id="c1897-110">Note that `find` throws an `ItemNotFound` error if the specified string doesn't exist in the range.</span></span> <span data-ttu-id="c1897-111">Si vous pensez que la chaîne spécifiée peut ne pas exister dans la plage, utilisez la méthode[findAllOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) à la place, pour que votre code gère ce scénario plus facilement.</span><span class="sxs-lookup"><span data-stu-id="c1897-111">If you expect that the specified string may not exist in the range, use the [findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) method instead, so your code gracefully handles that scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var table = sheet.tables.getItem("ExpensesTable");
    var searchRange = table.getRange();
    var foundRange = searchRange.find("Food", {
        completeMatch: true, // find will match the whole cell value
        matchCase: false, // find will not match case
        searchDirection: Excel.SearchDirection.forward // find will start searching at the beginning of the range
    });

    foundRange.load("address");
    return context.sync()
        .then(function() {
            console.log(foundRange.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="c1897-112">Lorsque la méthode `find` est appelée sur une plage représentant une cellule simple, la feuille de calcul entière est recherchée.</span><span class="sxs-lookup"><span data-stu-id="c1897-112">When the `find` method is called on a range representing a single cell, the entire worksheet is searched.</span></span> <span data-ttu-id="c1897-113">La recherche commence à cette cellule et continue dans la direction spécifiée par `SearchCriteria.searchDirection`, revenant à la ligne à la fin de la feuille de calcul si nécessaire.</span><span class="sxs-lookup"><span data-stu-id="c1897-113">The search begins at that cell and goes in the direction specified by `SearchCriteria.searchDirection`, wrapping around the ends of the worksheet if needed.</span></span>

## <a name="see-also"></a><span data-ttu-id="c1897-114">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c1897-114">See also</span></span>

- [<span data-ttu-id="c1897-115">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="c1897-115">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="c1897-116">Utiliser des cellules à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="c1897-116">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="c1897-117">Rechercher des cellules spéciales dans une plage à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="c1897-117">Find special cells within a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-special-cells.md)
