---
title: Effacer ou supprimer des plages à l’aide de l’API JavaScript pour Excel
description: Découvrez comment effacer ou supprimer des plages à l’aide de l’API JavaScript pour Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 7e030c6b5ba7ba6e6c54e9be0524cd93c2516bcb
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652869"
---
# <a name="clear-or-delete-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="73221-103">Effacer ou supprimer des plages à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="73221-103">Clear or delete ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="73221-104">Cet article fournit des exemples de code qui effacent et suppriment des plages avec l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="73221-104">This article provides code samples that clear and delete ranges with the Excel JavaScript API.</span></span> <span data-ttu-id="73221-105">Pour obtenir la liste complète des propriétés et des méthodes pris en charge par l’objet, voir `Range` [la classe Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="73221-105">For the complete list of properties and methods supported by the `Range` object, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="clear-a-range-of-cells"></a><span data-ttu-id="73221-106">Effacer une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="73221-106">Clear a range of cells</span></span>

<span data-ttu-id="73221-107">L’exemple de code suivant efface tout le contenu et la mise en forme des cellules de la plage **E2 : E5**.</span><span class="sxs-lookup"><span data-stu-id="73221-107">The following code sample clears all contents and formatting of cells in the range **E2:E5**.</span></span>  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-cleared"></a><span data-ttu-id="73221-108">Données avant l’effacement de la plage</span><span class="sxs-lookup"><span data-stu-id="73221-108">Data before range is cleared</span></span>

![Données dans Excel avant l’effacement de la plage](../images/excel-ranges-start.png)

### <a name="data-after-range-is-cleared"></a><span data-ttu-id="73221-110">Données après l’effacement de plage</span><span class="sxs-lookup"><span data-stu-id="73221-110">Data after range is cleared</span></span>

![Données dans Excel après l’effacement de plage](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a><span data-ttu-id="73221-112">Supprimer une plage de cellules</span><span class="sxs-lookup"><span data-stu-id="73221-112">Delete a range of cells</span></span>

<span data-ttu-id="73221-113">L’exemple de code suivant supprime les cellules de la plage **B4:E4** et déplace les autres cellules vers le haut pour remplir l’espace qui a été libéré par les cellules supprimées.</span><span class="sxs-lookup"><span data-stu-id="73221-113">The following code sample deletes the cells in the range **B4:E4** and shifts other cells up to fill the space that was vacated by the deleted cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-deleted"></a><span data-ttu-id="73221-114">Données avant la suppression d’une plage</span><span class="sxs-lookup"><span data-stu-id="73221-114">Data before range is deleted</span></span>

![Données dans Excel avant la suppression d’une plage](../images/excel-ranges-start.png)

### <a name="data-after-range-is-deleted"></a><span data-ttu-id="73221-116">Données après la suppression d’une plage</span><span class="sxs-lookup"><span data-stu-id="73221-116">Data after range is deleted</span></span>

![Données dans Excel après la suppression d’une plage](../images/excel-ranges-after-delete.png)


## <a name="see-also"></a><span data-ttu-id="73221-118">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="73221-118">See also</span></span>

- [<span data-ttu-id="73221-119">Utiliser des cellules à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="73221-119">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="73221-120">Définir et obtenir des plages à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="73221-120">Set and get ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get.md)
- [<span data-ttu-id="73221-121">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="73221-121">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)