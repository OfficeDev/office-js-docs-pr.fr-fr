---
title: Obtenir une plage à l’aide de Excel API JavaScript
description: Découvrez comment récupérer une plage à l’aide de l Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6aa9bb00bc9d24aeee5f1fef9e8d1531525e9d1f
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937285"
---
# <a name="get-a-range-using-the-excel-javascript-api"></a>Obtenir une plage à l’aide de Excel API JavaScript

Cet article fournit des exemples qui montrent différentes façons d’obtenir une plage dans une feuille de calcul à l’aide Excel API JavaScript. Pour obtenir la liste complète des propriétés et méthodes que l’objet prend en `Range` charge, [voir Excel. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="get-range-by-address"></a>Obtenir une plage en fonction d’une adresse

L’exemple de code suivant obtient la plage avec l’adresse **B2:C5** à partir de la feuille de calcul nommée **Sample,** charge sa propriété et écrit un `address` message dans la console.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:C5");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range B2:C5 is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-range-by-name"></a>Obtenir une plage en fonction d’un nom

L’exemple de code suivant obtient la plage nommée à partir de la feuille de calcul nommée Sample, charge sa propriété et écrit `MyRange` un message dans la  `address` console.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("MyRange");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range "MyRange" is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-used-range"></a>Obtenir une plage utilisée

L’exemple de code suivant obtient la plage utilisée à partir de la feuille de calcul nommée **Sample,** charge sa propriété et écrit `address` un message dans la console. La plage utilisée est la plus petite plage qui englobe toutes les cellules de la feuille de calcul auxquelles une valeur ou un format est affecté. Si la feuille de calcul entière est vide, la méthode renvoie une plage qui se compose uniquement de `getUsedRange()` la cellule supérieure gauche.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getUsedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the used range in the worksheet is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-entire-range"></a>Obtenir l’intégralité d’une plage

L’exemple de code suivant obtient l’ensemble de la plage de feuille de calcul à partir de la feuille de calcul nommée **Sample,** charge sa propriété et écrit un `address` message dans la console.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the entire worksheet range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide Excel API JavaScript](excel-add-ins-cells.md)
- [Insérer une plage à l’aide de Excel API JavaScript](excel-add-ins-ranges-insert.md)
