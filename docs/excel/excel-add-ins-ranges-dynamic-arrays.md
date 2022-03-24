---
title: Gérer les tableaux dynamiques et la plage qui se débordent à l’aide de Excel API JavaScript
description: Découvrez comment gérer les tableaux dynamiques et la plage qui se débordent avec l Excel API JavaScript.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 4ba4ab2bbce04465bc7db0a75e8ce39a6584a5a8
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745067"
---
# <a name="handle-dynamic-arrays-and-spilling-using-the-excel-javascript-api"></a>Gérer les tableaux dynamiques et les débordements à l’aide Excel API JavaScript

Cet article fournit un exemple de code qui gère les tableaux dynamiques et les étendues à l’aide de l Excel API JavaScript. Pour obtenir la liste complète des propriétés et méthodes que `Range` l’objet prend en charge, [voir Excel. Classe Range](/javascript/api/excel/excel.range).

## <a name="dynamic-arrays"></a>Tableaux dynamiques

Certaines Excel formules renvoyant [des tableaux dynamiques](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531). Ceux-ci remplissent les valeurs de plusieurs cellules en dehors de la cellule d’origine de la formule. Cette valeur de dépassement est appelée « dépassement ». Votre add-in peut trouver la plage utilisée pour un débordement avec la [méthode Range.getSpillingToRange](/javascript/api/excel/excel.range#excel-excel-range-getspillingtorange-member(1)) . Il existe également [une version *OrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties). `Range.getSpillingToRangeOrNullObject`

L’exemple suivant montre une formule de base qui copie le contenu d’une plage dans une cellule, qui se déborde dans les cellules voisines. Le add-in enregistre ensuite la plage qui contient le débordement.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    // Set G4 to a formula that returns a dynamic array.
    let targetCell = sheet.getRange("G4");
    targetCell.formulas = [["=A4:D4"]];

    // Get the address of the cells that the dynamic array spilled into.
    let spillRange = targetCell.getSpillingToRange();
    spillRange.load("address");

    // Sync and log the spilled-to range.
    await context.sync();

    // This will log the range as "G4:J4".
    console.log(`Copying the table headers spilled into ${spillRange.address}.`);
});
```

## <a name="range-spilling"></a>Étendue de plage

Recherchez la cellule responsable du débordement dans une cellule donnée à l’aide de la [méthode Range.getSpillParent](/javascript/api/excel/excel.range#excel-excel-range-getspillparent-member(1)) . Notez que `getSpillParent` fonctionne uniquement lorsque l’objet de plage est une seule cellule. L’appel `getSpillParent` sur une plage avec plusieurs cellules entraîne une erreur en cours de thrown (ou une plage null renvoyée pour `Range.getSpillParentOrNullObject`).

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide Excel API JavaScript](excel-add-ins-cells.md)
- [Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md)
