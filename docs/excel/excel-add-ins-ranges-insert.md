---
title: Insérer des plages à l’aide de Excel API JavaScript
description: Découvrez comment insérer une plage de cellules à l’Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0571e7d6140f5023008654a1e74d7abf6b3cab0a
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075781"
---
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a><span data-ttu-id="0dae6-103">Insérer une plage de cellules à l’aide de Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="0dae6-103">Insert a range of cells using the Excel JavaScript API</span></span>

<span data-ttu-id="0dae6-104">Cet article fournit un exemple de code qui insère une plage de cellules avec l Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="0dae6-104">This article provides a code sample that inserts a range of cells with the Excel JavaScript API.</span></span> <span data-ttu-id="0dae6-105">Pour obtenir la liste complète des propriétés et méthodes que l’objet prend en charge, voir `Range` [la Excel. Classe Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="0dae6-105">For the complete list of properties and methods that the `Range` object supports, see the [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="insert-a-range-of-cells"></a><span data-ttu-id="0dae6-106">Insérer une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="0dae6-106">Insert a range of cells</span></span>

<span data-ttu-id="0dae6-107">L’exemple de code suivant insère une plage de cellules dans l’emplacement **B4:E4** et déplace les autres cellules vers le bas pour laisser de l’espace pour les nouvelles cellules.</span><span class="sxs-lookup"><span data-stu-id="0dae6-107">The following code sample inserts a range of cells in location **B4:E4** and shifts other cells down to provide space for the new cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-inserted"></a><span data-ttu-id="0dae6-108">Données avant l’insertion de la plage</span><span class="sxs-lookup"><span data-stu-id="0dae6-108">Data before range is inserted</span></span>

![Données dans Excel avant l’insertion de la plage.](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a><span data-ttu-id="0dae6-110">Données après l’insertion de la plage</span><span class="sxs-lookup"><span data-stu-id="0dae6-110">Data after range is inserted</span></span>

![Données dans Excel après l’insertion de la plage.](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a><span data-ttu-id="0dae6-112">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="0dae6-112">See also</span></span>

- [<span data-ttu-id="0dae6-113">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="0dae6-113">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="0dae6-114">Utiliser des cellules à l’aide de Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="0dae6-114">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="0dae6-115">Effacer ou supprimer des plages à l’aide de l Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="0dae6-115">Clear or delete a ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-clear-delete.md)
