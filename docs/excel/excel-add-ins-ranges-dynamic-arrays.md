---
title: Gérer les tableaux dynamiques et la plage qui se débordent à l’aide de Excel API JavaScript
description: Découvrez comment gérer les tableaux dynamiques et la plage qui se débordent avec l Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 4601cd6a901243cadab0e7c5ead6061e28806377
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150484"
---
# <a name="handle-dynamic-arrays-and-spilling-using-the-excel-javascript-api"></a>Gérer les tableaux dynamiques et les débordements à l’aide Excel API JavaScript

Cet article fournit un exemple de code qui gère les tableaux dynamiques et les étendues à l’aide de l Excel API JavaScript. Pour obtenir la liste complète des propriétés et méthodes que l’objet prend en `Range` charge, [voir Excel. Classe Range](/javascript/api/excel/excel.range).

## <a name="dynamic-arrays"></a>Tableaux dynamiques

Certaines Excel formules de retour [dynamiques](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531). Ceux-ci remplissent les valeurs de plusieurs cellules en dehors de la cellule d’origine de la formule. Cette valeur de dépassement est appelée « débordement ». Votre add-in peut trouver la plage utilisée pour un débordement avec la [méthode Range.getSpillingToRange.](/javascript/api/excel/excel.range#getSpillingToRange__) Il existe également [une version *OrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties), `Range.getSpillingToRangeOrNullObject` .

L’exemple suivant montre une formule de base qui copie le contenu d’une plage dans une cellule, qui se renverse dans les cellules voisines. Le add-in enregistre ensuite la plage qui contient le débordement.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Set G4 to a formula that returns a dynamic array.
    var targetCell = sheet.getRange("G4");
    targetCell.formulas = [["=A4:D4"]];

    // Get the address of the cells that the dynamic array spilled into.
    var spillRange = targetCell.getSpillingToRange();
    spillRange.load("address");

    // Sync and log the spilled-to range.
    return context.sync().then(function () {
        // This will log the range as "G4:J4".
        console.log(`Copying the table headers spilled into ${spillRange.address}.`);
    });
}).catch(errorHandlerFunction);
```

## <a name="range-spilling"></a>Étendue de plage

Recherchez la cellule responsable du débordement dans une cellule donnée à l’aide de la [méthode Range.getSpillParent.](/javascript/api/excel/excel.range#getSpillParent__) Notez que `getSpillParent` fonctionne uniquement lorsque l’objet de plage est une seule cellule. L’appel sur une plage avec plusieurs cellules entraîne une erreur en cours de thrown (ou une plage `getSpillParent` null renvoyée pour `Range.getSpillParentOrNullObject` ).

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide Excel API JavaScript](excel-add-ins-cells.md)
- [Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md)
