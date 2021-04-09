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
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a>Insérer une plage de cellules à l’aide de l’API JavaScript pour Excel

Cet article fournit un exemple de code qui insère une plage de cellules avec l’API JavaScript pour Excel. Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` la [classe Excel.Range.](/javascript/api/excel/excel.range)

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="insert-a-range-of-cells"></a>Insérer une plage de cellules

L’exemple de code suivant insère une plage de cellules dans l’emplacement **B4:E4** et déplace les autres cellules vers le bas pour laisser de l’espace pour les nouvelles cellules.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-inserted"></a>Données avant l’insertion de la plage

![Données dans Excel avant l’insertion de la plage](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a>Données après l’insertion de la plage

![Données dans Excel après l’insertion de plage](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide de l’API JavaScript pour Excel](excel-add-ins-cells.md)
- [Effacer ou supprimer des plages à l’aide de l’API JavaScript pour Excel](excel-add-ins-ranges-clear-delete.md)
