---
title: Définir et obtenir la plage sélectionnée à l’aide de l Excel API JavaScript
description: Découvrez comment utiliser l’API JavaScript Excel pour définir et obtenir la plage sélectionnée à l’aide de l Excel API JavaScript.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 9517c072fae92b1b541a52b1805834c2bb429dd3
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745361"
---
# <a name="set-and-get-the-selected-range-using-the-excel-javascript-api"></a>Définir et obtenir la plage sélectionnée à l’aide de l Excel API JavaScript

Cet article fournit des exemples de code qui définissent et obtiennent la plage sélectionnée avec Excel API JavaScript. Pour obtenir la liste complète des propriétés et méthodes que `Range` l’objet prend en charge, [voir Excel. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a>Définir la plage sélectionnée

L’exemple de code suivant sélectionne la plage **B2:E6** dans la feuille de calcul active.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("B2:E6");

    range.select();

    await context.sync();
});
```

### <a name="selected-range-b2e6"></a>Plage sélectionnée  B2:E6

![Plage sélectionnée en Excel.](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a>Obtenir la plage sélectionnée

L’exemple de code suivant obtient la plage sélectionnée, charge sa `address` propriété et écrit un message dans la console.

```js
await Excel.run(async (context) => {
    let range = context.workbook.getSelectedRange();
    range.load("address");

    await context.sync();
    
    console.log(`The address of the selected range is "${range.address}"`);
});
```

## <a name="select-the-edge-of-a-used-range"></a>Sélectionner le bord d’une plage utilisée

Les méthodes [Range.getRangeEdge](/javascript/api/excel/excel.range#excel-excel-range-getrangeedge-member(1)) et [Range.getExtendedRange](/javascript/api/excel/excel.range#excel-excel-range-getextendedrange-member(1)) permet à votre add-in de répliquer le comportement des raccourcis de sélection du clavier, en sélectionnant le bord de la plage utilisée en fonction de la plage actuellement sélectionnée. Pour en savoir plus sur les plages utilisées, voir [Obtenir une plage utilisée](excel-add-ins-ranges-get.md#get-used-range).

Dans la capture d’écran suivante, la plage utilisée est le tableau avec des valeurs dans chaque cellule, **C5:F12**. Les cellules vides en dehors de ce tableau sont en dehors de la plage utilisée.

![Tableau avec des données de C5:F12 Excel.](../images/excel-ranges-used-range.png)

### <a name="select-the-cell-at-the-edge-of-the-current-used-range"></a>Sélectionner la cellule au bord de la plage utilisée actuelle

L’exemple de `Range.getRangeEdge` code suivant montre comment utiliser la méthode pour sélectionner la cellule au bord le plus proche de la plage utilisée actuelle, dans la direction vers le haut. Cette action correspond au résultat de l’utilisation du raccourci clavier de touche fléchée Ctrl+Haut pendant qu’une plage est sélectionnée.

```js
await Excel.run(async (context) => {
    // Get the selected range.
    let range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    let direction = Excel.KeyboardDirection.up;

    // Get the active cell in the workbook.
    let activeCell = context.workbook.getActiveCell();

    // Get the top-most cell of the current used range.
    // This method acts like the Ctrl+Up arrow key keyboard shortcut while a range is selected.
    let rangeEdge = range.getRangeEdge(
      direction,
      activeCell
    );
    rangeEdge.select();

    await context.sync();
});
```

#### <a name="before-selecting-the-cell-at-the-edge-of-the-used-range"></a>Avant de sélectionner la cellule au bord de la plage utilisée

La capture d’écran suivante montre une plage utilisée et une plage sélectionnée dans la plage utilisée. La plage utilisée est un tableau avec des données **au niveau de C5:F12**. Dans ce tableau, la plage **D8:E9** est sélectionnée. Cette sélection est à *l’état* antérieur, avant l’exécution de la `Range.getRangeEdge` méthode.

![Tableau avec des données de C5:F12 Excel. La plage D8:E9 est sélectionnée.](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-the-cell-at-the-edge-of-the-used-range"></a>Après avoir sélectionné la cellule au bord de la plage utilisée

La capture d’écran suivante montre le même tableau que la capture d’écran précédente, avec des données de la plage **C5:F12**. Dans ce tableau, la plage **D5** est sélectionnée. Cette sélection est *après l’état* , après `Range.getRangeEdge` l’exécution de la méthode pour sélectionner la cellule au bord de la plage utilisée dans la direction vers le haut.

![Tableau avec des données de C5:F12 Excel. La plage D5 est sélectionnée.](../images/excel-ranges-used-range-d5.png)

### <a name="select-all-cells-from-current-range-to-furthest-edge-of-used-range"></a>Sélectionner toutes les cellules de la plage actuelle au bord le plus proche de la plage utilisée

L’exemple de code `Range.getExtendedRange` suivant montre comment utiliser la méthode pour sélectionner toutes les cellules de la plage actuellement sélectionnée jusqu’au bord le plus proche de la plage utilisée, dans la direction vers le bas. Cette action correspond au résultat de l’utilisation du raccourci clavier avec touches de direction Ctrl+Shift+Bas pendant qu’une plage est sélectionnée.

```js
await Excel.run(async (context) => {
    // Get the selected range.
    let range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    let direction = Excel.KeyboardDirection.down;

    // Get the active cell in the workbook.
    let activeCell = context.workbook.getActiveCell();

    // Get all the cells from the currently selected range to the bottom-most edge of the used range.
    // This method acts like the Ctrl+Shift+Down arrow key keyboard shortcut while a range is selected.
    let extendedRange = range.getExtendedRange(
      direction,
      activeCell
    );
    extendedRange.select();

    await context.sync();
});
```

#### <a name="before-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a>Avant de sélectionner toutes les cellules de la plage actuelle au bord de la plage utilisée

La capture d’écran suivante montre une plage utilisée et une plage sélectionnée dans la plage utilisée. La plage utilisée est un tableau avec des données **au niveau de C5:F12**. Dans ce tableau, la plage **D8:E9** est sélectionnée. Cette sélection est à *l’état* antérieur, avant l’exécution de la `Range.getExtendedRange` méthode.

![Tableau avec des données de C5:F12 Excel. La plage D8:E9 est sélectionnée.](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a>Après avoir sélectionné toutes les cellules de la plage actuelle au bord de la plage utilisée

La capture d’écran suivante montre le même tableau que la capture d’écran précédente, avec des données de la plage **C5:F12**. Dans ce tableau, la plage **D8:E12** est sélectionnée. Cette sélection *s’exécute* après l’exécution `Range.getExtendedRange` de la méthode pour sélectionner toutes les cellules de la plage actuelle au bord de la plage utilisée dans la direction vers le bas.

![Tableau avec des données de C5:F12 Excel. La plage D8:E12 est sélectionnée.](../images/excel-ranges-used-range-d8-e12.png)

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide Excel API JavaScript](excel-add-ins-cells.md)
- [Définir et obtenir des valeurs de plage, du texte ou des formules à l’aide Excel API JavaScript](excel-add-ins-ranges-set-get-values.md)
- [Définir le format de plage à l’aide Excel API JavaScript](excel-add-ins-ranges-set-format.md)
