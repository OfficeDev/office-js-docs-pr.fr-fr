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
# <a name="set-range-format-using-the-excel-javascript-api"></a>Définir le format de plage à l’aide de l’API JavaScript pour Excel

Cet article fournit des exemples de code qui définissent la couleur de police, la couleur de remplissage et le format de nombre pour les cellules d’une plage avec l’API JavaScript pour Excel. Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` la classe [Excel.Range.](/javascript/api/excel/excel.range)

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-font-color-and-fill-color"></a>Définir la couleur de police et la couleur de remplissage

L’exemple de code ci-dessous définit la couleur de police et la couleur de remplissage des cellules de la plage **B2:E2**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-in-range-before-font-color-and-fill-color-are-set"></a>Données de la plage avant la définition de la couleur de police et de la couleur de remplissage

![Données dans Excel de la plage avant la définition de la couleur de police et de la couleur de remplissage](../images/excel-ranges-format-before.png)

### <a name="data-in-range-after-font-color-and-fill-color-are-set"></a>Données de la plage après la définition de la couleur de police et de la couleur de remplissage

![Données dans Excel après la définition du format](../images/excel-ranges-format-font-and-fill.png)

## <a name="set-number-format"></a>Définir le format de nombre

L’exemple de code ci-dessous définit le format de nombre des cellules dans la plage **D3:E5**.

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

### <a name="data-in-range-before-number-format-is-set"></a>Données de la plage avant la définition du format de nombre

![Données dans Excel avant la mise en forme des nombres](../images/excel-ranges-format-font-and-fill.png)

### <a name="data-in-range-after-number-format-is-set"></a>Données de la plage après la définition du format de nombre

![Données dans Excel après la mise en forme des nombres](../images/excel-ranges-format-numbers.png)

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide de l’API JavaScript pour Excel](excel-add-ins-cells.md)
- [Définir et obtenir des plages à l’aide de l’API JavaScript pour Excel](excel-add-ins-ranges-set-get.md)
- [Définir et obtenir des valeurs de plage, du texte ou des formules à l’aide de l’API JavaScript pour Excel](excel-add-ins-ranges-set-get-values.md)
