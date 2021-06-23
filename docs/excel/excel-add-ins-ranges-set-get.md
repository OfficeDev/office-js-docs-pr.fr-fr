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
# <a name="set-and-get-ranges-using-the-excel-javascript-api"></a>Définir et obtenir des plages à l’aide de Excel API JavaScript

Cet article fournit des exemples de code qui définissent et obtiennent des plages avec Excel API JavaScript. Pour obtenir la liste complète des propriétés et méthodes que l’objet prend en `Range` charge, [voir Excel. Classe Range](/javascript/api/excel/excel.range).

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

![Plage sélectionnée en Excel.](../images/excel-ranges-set-selection.png)

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
- [Utiliser des cellules à l’aide de Excel API JavaScript](excel-add-ins-cells.md)
- [Définir et obtenir des valeurs de plage, du texte ou des formules à l’aide Excel API JavaScript](excel-add-ins-ranges-set-get-values.md)
- [Définir le format de plage à l’aide Excel API JavaScript](excel-add-ins-ranges-set-format.md)
