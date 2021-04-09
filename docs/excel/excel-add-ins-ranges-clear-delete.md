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
# <a name="clear-or-delete-ranges-using-the-excel-javascript-api"></a>Effacer ou supprimer des plages à l’aide de l’API JavaScript pour Excel

Cet article fournit des exemples de code qui effacent et suppriment des plages avec l’API JavaScript pour Excel. Pour obtenir la liste complète des propriétés et des méthodes pris en charge par l’objet, voir `Range` [la classe Excel.Range.](/javascript/api/excel/excel.range)

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

![Données dans Excel avant l’effacement de la plage](../images/excel-ranges-start.png)

### <a name="data-after-range-is-cleared"></a>Données après l’effacement de plage

![Données dans Excel après l’effacement de plage](../images/excel-ranges-after-clear.png)

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

![Données dans Excel avant la suppression d’une plage](../images/excel-ranges-start.png)

### <a name="data-after-range-is-deleted"></a>Données après la suppression d’une plage

![Données dans Excel après la suppression d’une plage](../images/excel-ranges-after-delete.png)


## <a name="see-also"></a>Voir aussi

- [Utiliser des cellules à l’aide de l’API JavaScript pour Excel](excel-add-ins-cells.md)
- [Définir et obtenir des plages à l’aide de l’API JavaScript pour Excel](excel-add-ins-ranges-set-get.md)
- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
