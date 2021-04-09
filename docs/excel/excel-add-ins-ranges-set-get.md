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
# <a name="set-and-get-ranges-using-the-excel-javascript-api"></a>Définir et obtenir des plages à l’aide de l’API JavaScript pour Excel

Cet article fournit des exemples de code qui définissent et obtiennent des plages avec l’API JavaScript pour Excel. Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` la classe [Excel.Range.](/javascript/api/excel/excel.range)

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a>Définir la plage sélectionnée

L’exemple de code suivant sélectionne la plage **B2:E6** dans la feuille de calcul active.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a>Plage sélectionnée  B2:E6

![Plage sélectionnée dans Excel](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a>Obtenir la plage sélectionnée

L’exemple de code suivant obtient la plage sélectionnée, charge sa propriété et écrit `address` un message dans la console.

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

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide de l’API JavaScript pour Excel](excel-add-ins-cells.md)
- [Définir et obtenir des valeurs de plage, du texte ou des formules à l’aide de l’API JavaScript pour Excel](excel-add-ins-ranges-set-get-values.md)
- [Définir le format de plage à l’aide de l’API JavaScript pour Excel](excel-add-ins-ranges-set-format.md)
