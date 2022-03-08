---
title: Obtenir une plage à l’aide de Excel API JavaScript
description: Découvrez comment récupérer une plage à l’aide de l Excel API JavaScript.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 16c42ccf8f3496316fbf7b52e4d8139f819c6da1
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340938"
---
# <a name="get-a-range-using-the-excel-javascript-api"></a>Obtenir une plage à l’aide de Excel API JavaScript

Cet article fournit des exemples qui montrent différentes façons d’obtenir une plage dans une feuille de calcul à l’aide Excel API JavaScript. Pour obtenir la liste complète des propriétés et méthodes que `Range` l’objet prend en charge, [voir Excel. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="get-range-by-address"></a>Obtenir une plage en fonction d’une adresse

L’exemple de code suivant obtient la plage avec l’adresse **B2:C5** à partir de la feuille de calcul nommée **Sample**, `address` charge sa propriété et écrit un message dans la console.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    
    let range = sheet.getRange("B2:C5");
    range.load("address");
    await context.sync();
    
    console.log(`The address of the range B2:C5 is "${range.address}"`);
});
```

## <a name="get-range-by-name"></a>Obtenir une plage en fonction d’un nom

L’exemple de code suivant obtient la plage `MyRange` nommée à partir de la feuille de calcul nommée **Sample**, charge sa `address` propriété et écrit un message dans la console.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("MyRange");
    range.load("address");
    await context.sync();

    console.log(`The address of the range "MyRange" is "${range.address}"`);
});
```

## <a name="get-used-range"></a>Obtenir une plage utilisée

L’exemple de code suivant obtient la plage utilisée à partir de la feuille de calcul nommée **Sample**, charge sa `address` propriété et écrit un message dans la console. La plage utilisée est la plus petite plage qui englobe toutes les cellules de la feuille de calcul auxquelles une valeur ou un format est affecté. Si la feuille de calcul entière est vide, la `getUsedRange()` méthode renvoie une plage qui se compose uniquement de la cellule supérieure gauche.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getUsedRange();
    range.load("address");
    await context.sync();
    
    console.log(`The address of the used range in the worksheet is "${range.address}"`);
});
```

## <a name="get-entire-range"></a>Obtenir l’intégralité d’une plage

L’exemple de code suivant obtient l’ensemble de la plage de feuille de calcul à partir de la feuille de calcul nommée **Sample**, charge sa `address` propriété et écrit un message dans la console.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange();
    range.load("address");
    await context.sync();
    
    console.log(`The address of the entire worksheet range is "${range.address}"`);
});
```

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide Excel API JavaScript](excel-add-ins-cells.md)
- [Insérer une plage à l’aide de Excel API JavaScript](excel-add-ins-ranges-insert.md)
