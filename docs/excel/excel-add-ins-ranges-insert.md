---
title: Insérer des plages à l’aide Excel API JavaScript
description: Découvrez comment insérer une plage de cellules avec l’API Excel JavaScript.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: ce559d0726b7d69c5f4c8c6d00a4e714c04df735
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745212"
---
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a>Insérer une plage de cellules à l’aide de l Excel API JavaScript

Cet article fournit un exemple de code qui insère une plage de cellules avec l Excel API JavaScript. Pour obtenir la liste complète des propriétés et méthodes que `Range` l’objet prend en charge, voir [la Excel. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="insert-a-range-of-cells"></a>Insérer une plage de cellules

L’exemple de code suivant insère une plage de cellules dans l’emplacement **B4:E4** et déplace les autres cellules vers le bas pour laisser de l’espace pour les nouvelles cellules.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    await context.sync();
});
```

### <a name="data-before-range-is-inserted"></a>Données avant l’insertion de la plage

![Données dans Excel avant l’insertion de la plage.](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a>Données après l’insertion de la plage

![Données dans Excel après l’insertion de la plage.](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide Excel API JavaScript](excel-add-ins-cells.md)
- [Effacer ou supprimer des plages à l’aide de l Excel API JavaScript](excel-add-ins-ranges-clear-delete.md)
