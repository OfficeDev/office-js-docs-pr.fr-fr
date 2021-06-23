---
title: Définir et obtenir la plage sélectionnée à l’aide de l Excel API JavaScript
description: Découvrez comment utiliser l’API JavaScript Excel pour définir et obtenir des plages à l’aide de l Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0bd4a4f4bcf40e7899ee429cdc631a43ba176077
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075774"
---
# <a name="set-and-get-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="51a1c-103">Définir et obtenir des plages à l’aide de Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="51a1c-103">Set and get ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="51a1c-104">Cet article fournit des exemples de code qui définissent et obtiennent des plages avec Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="51a1c-104">This article provides code samples that set and get ranges with the Excel JavaScript API.</span></span> <span data-ttu-id="51a1c-105">Pour obtenir la liste complète des propriétés et méthodes que l’objet prend en `Range` charge, [voir Excel. Classe Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="51a1c-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a><span data-ttu-id="51a1c-106">Définir la plage sélectionnée</span><span class="sxs-lookup"><span data-stu-id="51a1c-106">Set the selected range</span></span>

<span data-ttu-id="51a1c-107">L’exemple de code suivant sélectionne la plage **B2:E6** dans la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="51a1c-107">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a><span data-ttu-id="51a1c-108">Plage sélectionnée  B2:E6</span><span class="sxs-lookup"><span data-stu-id="51a1c-108">Selected range B2:E6</span></span>

![Plage sélectionnée en Excel.](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="51a1c-110">Obtenir la plage sélectionnée</span><span class="sxs-lookup"><span data-stu-id="51a1c-110">Get the selected range</span></span>

<span data-ttu-id="51a1c-111">L’exemple de code suivant obtient la plage sélectionnée, charge sa propriété et écrit `address` un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="51a1c-111">The following code sample gets the selected range, loads its `address` property, and writes a message to the console.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="51a1c-112">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="51a1c-112">See also</span></span>

- [<span data-ttu-id="51a1c-113">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="51a1c-113">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="51a1c-114">Utiliser des cellules à l’aide de Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="51a1c-114">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="51a1c-115">Définir et obtenir des valeurs de plage, du texte ou des formules à l’aide Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="51a1c-115">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
- [<span data-ttu-id="51a1c-116">Définir le format de plage à l’aide Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="51a1c-116">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)
