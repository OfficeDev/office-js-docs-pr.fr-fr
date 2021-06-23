---
title: Effacer ou supprimer des plages à l’aide de Excel API JavaScript
description: Découvrez comment effacer ou supprimer des plages à l’aide de l Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a1bd99db3aa9af3903552d9cefc6ec6d21701136
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075830"
---
# <a name="clear-or-delete-ranges-using-the-excel-javascript-api"></a>Effacer ou supprimer des plages à l’aide de Excel API JavaScript

Cet article fournit des exemples de code qui effacent et suppriment des plages avec l Excel API JavaScript. Pour obtenir la liste complète des propriétés et méthodes pris en charge par `Range` [l’objet, voir Excel. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="clear-a-range-of-cells"></a>Effacer une plage de cellules

L’exemple de code suivant efface tout le contenu et la mise en forme des cellules de la plage **E2 : E5**.  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-cleared"></a>Données avant l’effacement de la plage

![Données dans Excel avant l’effacée de la plage.](../images/excel-ranges-start.png)

### <a name="data-after-range-is-cleared"></a>Données après l’effacement de plage

![Données dans Excel une fois la plage effacée.](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a>Supprimer une plage de cellules

L’exemple de code suivant supprime les cellules de la plage **B4:E4** et déplace les autres cellules vers le haut pour remplir l’espace qui a été libéré par les cellules supprimées.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-deleted"></a>Données avant la suppression d’une plage

![Données dans Excel avant la suppression de la plage.](../images/excel-ranges-start.png)

### <a name="data-after-range-is-deleted"></a>Données après la suppression d’une plage

![Données dans Excel une fois la plage supprimée.](../images/excel-ranges-after-delete.png)


## <a name="see-also"></a>Voir aussi

- [Utiliser des cellules à l’aide de Excel API JavaScript](excel-add-ins-cells.md)
- [Définir et obtenir des plages à l’aide de Excel API JavaScript](excel-add-ins-ranges-set-get.md)
- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
