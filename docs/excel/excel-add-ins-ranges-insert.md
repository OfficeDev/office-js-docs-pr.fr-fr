---
title: Insérer des plages à l’aide de l’API JavaScript pour Excel
description: Découvrez comment insérer une plage de cellules avec l’API JavaScript pour Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 401a08dd10b3775012738ab9c80ec6ab367555ec
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652844"
---
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a><span data-ttu-id="6c33a-103">Insérer une plage de cellules à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="6c33a-103">Insert a range of cells using the Excel JavaScript API</span></span>

<span data-ttu-id="6c33a-104">Cet article fournit un exemple de code qui insère une plage de cellules avec l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="6c33a-104">This article provides a code sample that inserts a range of cells with the Excel JavaScript API.</span></span> <span data-ttu-id="6c33a-105">Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` la [classe Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="6c33a-105">For the complete list of properties and methods that the `Range` object supports, see the [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="insert-a-range-of-cells"></a><span data-ttu-id="6c33a-106">Insérer une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="6c33a-106">Insert a range of cells</span></span>

<span data-ttu-id="6c33a-107">L’exemple de code suivant insère une plage de cellules dans l’emplacement **B4:E4** et déplace les autres cellules vers le bas pour laisser de l’espace pour les nouvelles cellules.</span><span class="sxs-lookup"><span data-stu-id="6c33a-107">The following code sample inserts a range of cells in location **B4:E4** and shifts other cells down to provide space for the new cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-inserted"></a><span data-ttu-id="6c33a-108">Données avant l’insertion de la plage</span><span class="sxs-lookup"><span data-stu-id="6c33a-108">Data before range is inserted</span></span>

![Données dans Excel avant l’insertion de la plage](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a><span data-ttu-id="6c33a-110">Données après l’insertion de la plage</span><span class="sxs-lookup"><span data-stu-id="6c33a-110">Data after range is inserted</span></span>

![Données dans Excel après l’insertion de plage](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a><span data-ttu-id="6c33a-112">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="6c33a-112">See also</span></span>

- [<span data-ttu-id="6c33a-113">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="6c33a-113">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="6c33a-114">Utiliser des cellules à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="6c33a-114">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="6c33a-115">Effacer ou supprimer des plages à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="6c33a-115">Clear or delete a ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-clear-delete.md)
