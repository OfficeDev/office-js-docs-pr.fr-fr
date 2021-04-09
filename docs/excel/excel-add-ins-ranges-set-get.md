---
title: Définir et obtenir la plage sélectionnée à l’aide de l’API JavaScript pour Excel
description: Découvrez comment utiliser l’API JavaScript excel pour définir et obtenir des plages à l’aide de l’API JavaScript pour Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 06b6219924f0667ecef57d608cb417a76ef8031d
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652817"
---
# <a name="set-and-get-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="077dc-103">Définir et obtenir des plages à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="077dc-103">Set and get ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="077dc-104">Cet article fournit des exemples de code qui définissent et obtiennent des plages avec l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="077dc-104">This article provides code samples that set and get ranges with the Excel JavaScript API.</span></span> <span data-ttu-id="077dc-105">Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` la classe [Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="077dc-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a><span data-ttu-id="077dc-106">Définir la plage sélectionnée</span><span class="sxs-lookup"><span data-stu-id="077dc-106">Set the selected range</span></span>

<span data-ttu-id="077dc-107">L’exemple de code suivant sélectionne la plage **B2:E6** dans la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="077dc-107">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a><span data-ttu-id="077dc-108">Plage sélectionnée  B2:E6</span><span class="sxs-lookup"><span data-stu-id="077dc-108">Selected range B2:E6</span></span>

![Plage sélectionnée dans Excel](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="077dc-110">Obtenir la plage sélectionnée</span><span class="sxs-lookup"><span data-stu-id="077dc-110">Get the selected range</span></span>

<span data-ttu-id="077dc-111">L’exemple de code suivant obtient la plage sélectionnée, charge sa propriété et écrit `address` un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="077dc-111">The following code sample gets the selected range, loads its `address` property, and writes a message to the console.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="077dc-112">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="077dc-112">See also</span></span>

- [<span data-ttu-id="077dc-113">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="077dc-113">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="077dc-114">Utiliser des cellules à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="077dc-114">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="077dc-115">Définir et obtenir des valeurs de plage, du texte ou des formules à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="077dc-115">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
- [<span data-ttu-id="077dc-116">Définir le format de plage à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="077dc-116">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)
