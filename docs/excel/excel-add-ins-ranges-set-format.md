---
title: Définir le format d’une plage à l’aide de l’API JavaScript pour Excel
description: Découvrez comment utiliser l’API JavaScript excel pour définir le format d’une plage.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: fdd78ea69fc38cbefb9d240dbc61554891c73c21
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652836"
---
# <a name="set-range-format-using-the-excel-javascript-api"></a><span data-ttu-id="7707f-103">Définir le format de plage à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="7707f-103">Set range format using the Excel JavaScript API</span></span>

<span data-ttu-id="7707f-104">Cet article fournit des exemples de code qui définissent la couleur de police, la couleur de remplissage et le format de nombre pour les cellules d’une plage avec l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="7707f-104">This article provides code samples that set font color, fill color, and number format for cells in a range with the Excel JavaScript API.</span></span> <span data-ttu-id="7707f-105">Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` la classe [Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="7707f-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-font-color-and-fill-color"></a><span data-ttu-id="7707f-106">Définir la couleur de police et la couleur de remplissage</span><span class="sxs-lookup"><span data-stu-id="7707f-106">Set font color and fill color</span></span>

<span data-ttu-id="7707f-107">L’exemple de code ci-dessous définit la couleur de police et la couleur de remplissage des cellules de la plage **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="7707f-107">The following code sample sets the font color and fill color for cells in range **B2:E2**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-in-range-before-font-color-and-fill-color-are-set"></a><span data-ttu-id="7707f-108">Données de la plage avant la définition de la couleur de police et de la couleur de remplissage</span><span class="sxs-lookup"><span data-stu-id="7707f-108">Data in range before font color and fill color are set</span></span>

![Données dans Excel de la plage avant la définition de la couleur de police et de la couleur de remplissage](../images/excel-ranges-format-before.png)

### <a name="data-in-range-after-font-color-and-fill-color-are-set"></a><span data-ttu-id="7707f-110">Données de la plage après la définition de la couleur de police et de la couleur de remplissage</span><span class="sxs-lookup"><span data-stu-id="7707f-110">Data in range after font color and fill color are set</span></span>

![Données dans Excel après la définition du format](../images/excel-ranges-format-font-and-fill.png)

## <a name="set-number-format"></a><span data-ttu-id="7707f-112">Définir le format de nombre</span><span class="sxs-lookup"><span data-stu-id="7707f-112">Set number format</span></span>

<span data-ttu-id="7707f-113">L’exemple de code ci-dessous définit le format de nombre des cellules dans la plage **D3:E5**.</span><span class="sxs-lookup"><span data-stu-id="7707f-113">The following code sample sets the number format for the cells in range **D3:E5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var formats = [
        ["0.00", "0.00"],
        ["0.00", "0.00"],
        ["0.00", "0.00"]
    ];

    var range = sheet.getRange("D3:E5");
    range.numberFormat = formats;

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-in-range-before-number-format-is-set"></a><span data-ttu-id="7707f-114">Données de la plage avant la définition du format de nombre</span><span class="sxs-lookup"><span data-stu-id="7707f-114">Data in range before number format is set</span></span>

![Données dans Excel avant la mise en forme des nombres](../images/excel-ranges-format-font-and-fill.png)

### <a name="data-in-range-after-number-format-is-set"></a><span data-ttu-id="7707f-116">Données de la plage après la définition du format de nombre</span><span class="sxs-lookup"><span data-stu-id="7707f-116">Data in range after number format is set</span></span>

![Données dans Excel après la mise en forme des nombres](../images/excel-ranges-format-numbers.png)

## <a name="see-also"></a><span data-ttu-id="7707f-118">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="7707f-118">See also</span></span>

- [<span data-ttu-id="7707f-119">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="7707f-119">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="7707f-120">Utiliser des cellules à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="7707f-120">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="7707f-121">Définir et obtenir des plages à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="7707f-121">Set and get ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get.md)
- [<span data-ttu-id="7707f-122">Définir et obtenir des valeurs de plage, du texte ou des formules à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="7707f-122">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)