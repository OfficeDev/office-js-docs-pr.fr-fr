---
title: Définir et obtenir la plage sélectionnée à l’aide de l Excel API JavaScript
description: Découvrez comment utiliser l’API JavaScript Excel pour définir et obtenir la plage sélectionnée à l’aide de Excel API JavaScript.
ms.date: 06/22/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9e4c31f165b39d45fac342cb85577ef737105472
ms.sourcegitcommit: ebb4a22a0bdeb5623c72b9494ebbce3909d0c90c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/25/2021
ms.locfileid: "53126726"
---
# <a name="set-and-get-the-selected-range-using-the-excel-javascript-api"></a>Définir et obtenir la plage sélectionnée à l’aide de l Excel API JavaScript

Cet article fournit des exemples de code qui définissent et obtiennent la plage sélectionnée avec Excel API JavaScript. Pour obtenir la liste complète des propriétés et méthodes que l’objet prend en `Range` charge, [voir Excel. Classe Range](/javascript/api/excel/excel.range).

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

## <a name="select-the-edge-of-a-used-range-online-only"></a>Sélectionner le bord d’une plage utilisée (en ligne uniquement)

> [!NOTE]
> Les `Range.getRangeEdge` méthodes et les méthodes sont actuellement disponibles uniquement dans `Range.getExtendedRange` ExcelApiOnline 1.1. Pour plus d’informations, voir Excel’ensemble de conditions requises de [l’API JavaScript en ligne uniquement.](../reference/requirement-sets/excel-api-online-requirement-set.md)

Les méthodes [Range.getRangeEdge](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_) et [Range.getExtendedRange](/javascript/api/excel/excel.range#getExtendedRange_directionString__activeCell_) vous permet de répliquer le comportement des raccourcis de sélection du clavier, en sélectionnant le bord de la plage utilisée en fonction de la plage actuellement sélectionnée. Pour en savoir plus sur les plages utilisées, voir [Obtenir une plage utilisée.](excel-add-ins-ranges-get.md#get-used-range)

Dans la capture d’écran suivante, la plage utilisée est le tableau avec des valeurs dans chaque cellule, **C5:F12**. Les cellules vides en dehors de ce tableau sont en dehors de la plage utilisée.

![Tableau avec des données de C5:F12 Excel.](../images/excel-ranges-used-range.png)

### <a name="select-the-cell-at-the-edge-of-the-current-used-range"></a>Sélectionner la cellule au bord de la plage utilisée actuelle

L’exemple de code suivant montre comment utiliser la méthode pour sélectionner la cellule au bord le plus proche de la plage utilisée actuelle, dans `Range.getRangeEdge` la direction vers le haut. Cette action correspond au résultat de l’utilisation du raccourci clavier de touche fléchée Ctrl+Haut pendant qu’une plage est sélectionnée.

```js
Excel.run(function (context) {
    // Get the selected range.
    var range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    var direction = Excel.KeyboardDirection.up;

    // Get the active cell in the workbook.
    var activeCell = context.workbook.getActiveCell();

    // Get the top-most cell of the current used range.
    // This method acts like the Ctrl+Up arrow key keyboard shortcut while a range is selected.
    var rangeEdge = range.getRangeEdge(
      direction,
      activeCell
    );
    rangeEdge.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="before-selecting-the-cell-at-the-edge-of-the-used-range"></a>Avant de sélectionner la cellule au bord de la plage utilisée

La capture d’écran suivante montre une plage utilisée et une plage sélectionnée dans la plage utilisée. La plage utilisée est un tableau avec des données **au niveau de C5:F12**. Dans ce tableau, la plage **D8:E9** est sélectionnée. Cette sélection est à *l’état* antérieur, avant l’exécution de la `Range.getRangeEdge` méthode.

![Tableau avec des données de C5:F12 Excel. La plage D8:E9 est sélectionnée.](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-the-cell-at-the-edge-of-the-used-range"></a>Après avoir sélectionné la cellule au bord de la plage utilisée

La capture d’écran suivante montre le même tableau que la capture d’écran précédente, avec des données de la plage **C5:F12**. Dans ce tableau, la plage **D5** est sélectionnée. Cette sélection *s’exécute après* l’exécution de la méthode pour sélectionner la cellule au bord de la plage utilisée dans la direction vers le `Range.getRangeEdge` haut.

![Tableau avec des données de C5:F12 Excel. La plage D5 est sélectionnée.](../images/excel-ranges-used-range-d5.png)

### <a name="select-all-cells-from-current-range-to-furthest-edge-of-used-range"></a>Sélectionner toutes les cellules de la plage actuelle au bord le plus proche de la plage utilisée

L’exemple de code suivant montre comment utiliser la méthode pour sélectionner toutes les cellules de la plage actuellement sélectionnée au bord le plus proche de la plage utilisée, dans la direction vers le `Range.getExtendedRange` bas. Cette action correspond au résultat de l’utilisation du raccourci clavier avec touches de direction Ctrl+Shift+Bas pendant qu’une plage est sélectionnée.

```js
Excel.run(function (context) {
    // Get the selected range.
    var range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    var direction = Excel.KeyboardDirection.down;

    // Get the active cell in the workbook.
    var activeCell = context.workbook.getActiveCell();

    // Get all the cells from the currently selected range to the bottom-most edge of the used range.
    // This method acts like the Ctrl+Shift+Down arrow key keyboard shortcut while a range is selected.
    var extendedRange = range.getExtendedRange(
      direction,
      activeCell
    );
    extendedRange.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="before-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a>Avant de sélectionner toutes les cellules de la plage actuelle au bord de la plage utilisée

La capture d’écran suivante montre une plage utilisée et une plage sélectionnée dans la plage utilisée. La plage utilisée est un tableau avec des données **au niveau de C5:F12**. Dans ce tableau, la plage **D8:E9** est sélectionnée. Cette sélection est à *l’état* antérieur, avant l’exécution de la `Range.getExtendedRange` méthode.

![Tableau avec des données de C5:F12 Excel. La plage D8:E9 est sélectionnée.](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a>Après avoir sélectionné toutes les cellules de la plage actuelle au bord de la plage utilisée

La capture d’écran suivante montre le même tableau que la capture d’écran précédente, avec des données de la plage **C5:F12**. Dans ce tableau, la plage **D8:E12** est sélectionnée. Cette sélection *s’exécute* après l’exécution de la méthode pour sélectionner toutes les cellules de la plage actuelle au bord de la plage utilisée dans `Range.getExtendedRange` la direction vers le bas.

![Tableau avec des données de C5:F12 Excel. La plage D8:E12 est sélectionnée.](../images/excel-ranges-used-range-d8-e12.png)

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide de Excel API JavaScript](excel-add-ins-cells.md)
- [Définir et obtenir des valeurs de plage, du texte ou des formules à l’aide Excel API JavaScript](excel-add-ins-ranges-set-get-values.md)
- [Définir le format de plage à l’aide Excel API JavaScript](excel-add-ins-ranges-set-format.md)
