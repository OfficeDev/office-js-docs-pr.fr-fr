---
title: Définir le format d’une plage à l’aide de Excel API JavaScript
description: Découvrez comment utiliser l’API JavaScript Excel pour définir le format d’une plage.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 41727f6fd71636be24bdc1bb8416cb3ba07c06e1
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340350"
---
# <a name="set-range-format-using-the-excel-javascript-api"></a>Définir le format de plage à l’aide Excel API JavaScript

Cet article fournit des exemples de code qui définissent la couleur de police, la couleur de remplissage et le format de nombre pour les cellules d’une plage avec Excel API JavaScript. Pour obtenir la liste complète des propriétés et méthodes que `Range` l’objet prend en charge, [voir Excel. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-font-color-and-fill-color"></a>Définir la couleur de police et la couleur de remplissage

L’exemple de code ci-dessous définit la couleur de police et la couleur de remplissage des cellules de la plage **B2:E2**.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";
    range.format.font.color = "white";

    await context.sync();
});
```

### <a name="data-in-range-before-font-color-and-fill-color-are-set"></a>Données de la plage avant la définition de la couleur de police et de la couleur de remplissage

![Données dans Excel avant la mise en forme.](../images/excel-ranges-format-before.png)

### <a name="data-in-range-after-font-color-and-fill-color-are-set"></a>Données de la plage après la définition de la couleur de police et de la couleur de remplissage

![Données en Excel après la mise en forme.](../images/excel-ranges-format-font-and-fill.png)

## <a name="set-number-format"></a>Définir le format de nombre

L’exemple de code ci-dessous définit le format de nombre des cellules dans la plage **D3:E5**.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let formats = [
        ["0.00", "0.00"],
        ["0.00", "0.00"],
        ["0.00", "0.00"]
    ];

    let range = sheet.getRange("D3:E5");
    range.numberFormat = formats;

    await context.sync();
});
```

### <a name="data-in-range-before-number-format-is-set"></a>Données de la plage avant la définition du format de nombre

![Données dans Excel avant la mise en forme du nombre.](../images/excel-ranges-format-font-and-fill.png)

### <a name="data-in-range-after-number-format-is-set"></a>Données de la plage après la définition du format de nombre

![Données dans Excel après la mise en forme du nombre.](../images/excel-ranges-format-numbers.png)

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide Excel API JavaScript](excel-add-ins-cells.md)
- [Définir et obtenir des plages à l’aide de Excel API JavaScript](excel-add-ins-ranges-set-get.md)
- [Définir et obtenir des valeurs de plage, du texte ou des formules à l’aide Excel API JavaScript](excel-add-ins-ranges-set-get-values.md)
