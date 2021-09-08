---
title: Définir et obtenir des valeurs de plage, du texte ou des formules à l’aide Excel API JavaScript
description: Découvrez comment utiliser l’API JavaScript Excel pour définir et obtenir des valeurs de plage, du texte ou des formules.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 5d9d1bf3b248585bf27ac591754cfa4eb4dd0fbc
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936588"
---
# <a name="set-and-get-range-values-text-or-formulas-using-the-excel-javascript-api"></a>Définir et obtenir des valeurs de plage, du texte ou des formules à l’aide Excel API JavaScript

Cet article fournit des exemples de code qui définissent et obtiennent des valeurs de plage, du texte ou des formules avec Excel API JavaScript. Pour obtenir la liste complète des propriétés et méthodes que l’objet prend en `Range` charge, [voir Excel. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-values-or-formulas"></a>Définir des valeurs ou des formules

Les exemples de code suivants définissent des valeurs et des formules pour une cellule unique ou une plage de cellules.

### <a name="set-value-for-a-single-cell"></a>Définir une valeur pour une cellule unique

L’exemple de code suivant définit la valeur de la cellule **C3** sur « 5 », puis définit la largeur des colonnes pour mieux s’adapter aux données.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-value-is-updated"></a>Données avant la mise à jour de la valeur de la cellule

![Données dans Excel avant la mise à jour de la valeur de la cellule.](../images/excel-ranges-set-start.png)

#### <a name="data-after-cell-value-is-updated"></a>Données après la mise à jour de la valeur de la cellule

![Données dans Excel une fois la valeur de la cellule mise à jour.](../images/excel-ranges-set-cell-value.png)

### <a name="set-values-for-a-range-of-cells"></a>Définir des valeurs pour une plage de cellules

L’exemple de code suivant définit les valeurs des cellules de la plage **B5:D5**, puis définit la largeur des colonnes pour mieux s’adapter aux données.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["Potato Chips", 10, 1.80],
    ];

    var range = sheet.getRange("B5:D5");
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-values-are-updated"></a>Données avant la mise à jour des valeurs des cellules

![Données dans les Excel avant la mise à jour des valeurs des cellules.](../images/excel-ranges-set-start.png)

#### <a name="data-after-cell-values-are-updated"></a>Données après la mise à jour des valeurs des cellules

![Données dans les Excel après la mise à jour des valeurs des cellules.](../images/excel-ranges-set-cell-values.png)

### <a name="set-formula-for-a-single-cell"></a>Définir la formule d’une cellule unique

L’exemple de code suivant définit une formule pour la cellule **E3**, puis définit la largeur des colonnes pour mieux s’adapter aux données.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-formula-is-set"></a>Données avant la définition de la formule de la cellule

![Données dans les Excel la formule de la cellule est définie.](../images/excel-ranges-start-set-formula.png)

#### <a name="data-after-cell-formula-is-set"></a>Données après la définition de la formule de la cellule

![Données dans les Excel une fois la formule de cellule définie.](../images/excel-ranges-set-formula.png)

### <a name="set-formulas-for-a-range-of-cells"></a>Définir des formules pour une plage de cellules

L’exemple de code ci-dessous définit des formules pour les cellules de la plage **E2:E6**, puis définit la largeur des colonnes pour mieux s’adapter aux données.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"]
    ];

    var range = sheet.getRange("E3:E6");
    range.formulas = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-formulas-are-set"></a>Données avant la définition des formules des cellules

![Données dans Excel avant la mise en place des formules de cellule.](../images/excel-ranges-start-set-formula.png)

#### <a name="data-after-cell-formulas-are-set"></a>Données après la définition des formules des cellules

![Données dans les Excel une fois que les formules de cellule sont définies.](../images/excel-ranges-set-formulas.png)

## <a name="get-values-text-or-formulas"></a>Obtenir des valeurs, du texte ou des formules

Ces exemples de code obtiennent des valeurs, du texte et des formules à partir d’une plage de cellules.

### <a name="get-values-from-a-range-of-cells"></a>Obtenir des valeurs à partir d’une plage de cellules

L’exemple de code suivant obtient la plage **B2:E6,** charge sa propriété et écrit les `values` valeurs dans la console. La `values` propriété d’une plage spécifie les valeurs brutes que contiennent les cellules. Même si certaines cellules d’une plage contiennent des formules, la propriété de la plage spécifie les valeurs brutes de ces cellules, et non l’une `values` des formules.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("values");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.values, null, 4));
        });
}).catch(errorHandlerFunction);
```

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a>Données de la plage (les valeurs de la colonne E sont le résultat des formules)

![Données dans les Excel une fois que les formules de cellule sont définies.](../images/excel-ranges-set-formulas.png)

#### <a name="rangevalues-as-logged-to-the-console-by-the-code-sample-above"></a>range.values (comme consigné dans la console par l’exemple de code ci-dessus)

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        15
    ],
    [
        "Coffee",
        1,
        34.5,
        34.5
    ],
    [
        "Chocolate",
        5,
        9.56,
        47.8
    ],
    [
        "",
        "",
        "",
        97.3
    ]
]
```

### <a name="get-text-from-a-range-of-cells"></a>Obtenir du texte à partir d’une plage de cellules

L’exemple de code suivant obtient la plage **B2:E6,** charge sa propriété `text` et l’écrit dans la console. La `text` propriété d’une plage spécifie les valeurs d’affichage des cellules de la plage. Même si certaines cellules d’une plage contiennent des formules, la propriété de la plage spécifie les valeurs d’affichage de ces cellules, et non des `text` formules.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("text");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.text, null, 4));
        });
}).catch(errorHandlerFunction);
```

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a>Données de la plage (les valeurs de la colonne E sont le résultat des formules)

![Données dans les Excel une fois que les formules de cellule sont définies.](../images/excel-ranges-set-formulas.png)

#### <a name="rangetext-as-logged-to-the-console-by-the-code-sample-above"></a>range.text (comme consigné dans la console par l’exemple de code ci-dessus)

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        "2",
        "7.5",
        "15"
    ],
    [
        "Coffee",
        "1",
        "34.5",
        "34.5"
    ],
    [
        "Chocolate",
        "5",
        "9.56",
        "47.8"
    ],
    [
        "",
        "",
        "",
        "97.3"
    ]
]
```

### <a name="get-formulas-from-a-range-of-cells"></a>Obtenir des formules à partir d’une plage de cellules

L’exemple de code suivant obtient la plage **B2:E6,** charge sa propriété `formulas` et l’écrit dans la console. La propriété d’une plage spécifie les formules des cellules de la plage qui contiennent des formules et les valeurs brutes des cellules de la plage qui ne contiennent pas `formulas` de formules.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("formulas");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.formulas, null, 4));
        });
}).catch(errorHandlerFunction);
```

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a>Données de la plage (les valeurs de la colonne E sont le résultat des formules)

![Données dans les Excel une fois que les formules de cellule sont définies.](../images/excel-ranges-set-formulas.png)

#### <a name="rangeformulas-as-logged-to-the-console-by-the-code-sample-above"></a>range.formulas (comme consigné dans la console par l’exemple de code ci-dessus)

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        "=C3 * D3"
    ],
    [
        "Coffee",
        1,
        34.5,
        "=C4 * D4"
    ],
    [
        "Chocolate",
        5,
        9.56,
        "=C5 * D5"
    ],
    [
        "",
        "",
        "",
        "=SUM(E3:E5)"
    ]
]
```

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide Excel API JavaScript](excel-add-ins-cells.md)
- [Définir et obtenir des plages à l’aide de Excel API JavaScript](excel-add-ins-ranges-set-get.md)
- [Définir le format de plage à l’aide Excel API JavaScript](excel-add-ins-ranges-set-format.md)
