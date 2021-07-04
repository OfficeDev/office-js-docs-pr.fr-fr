---
title: Utiliser des antécédents et des dépendances de formule à l’aide Excel API JavaScript
description: Découvrez comment utiliser l’API JavaScript Excel pour récupérer les antécédents et les dépendances de formule.
ms.date: 07/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: bf92400af00df42ac245b9a2d3ff5e72512b5722
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290774"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a>Obtenir des antécédents et des dépendances de formule à l’aide Excel API JavaScript

Excel formules font souvent référence à d’autres cellules. Ces références entre cellules sont appelées « antécédents » et « dépendants ». Un précédent est une cellule qui fournit des données à une formule. Une cellule dépendante est une cellule qui contient une formule qui fait référence à d’autres cellules. Pour en savoir plus sur Excel fonctionnalités liées aux relations entre les cellules, voir Afficher les relations entre les [formules et les cellules.](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507)

Une cellule peut avoir une cellule précédente et cette cellule peut avoir ses propres cellules précédentes. Un « précédent direct » est le premier groupe de cellules précédent dans cette séquence, similaire au concept de parents dans une relation parent-enfant. Un « dépendant direct » est le premier groupe dépendant de cellules dans une séquence, semblable aux enfants d’une relation parent-enfant. Les cellules qui font référence à d’autres cellules d’un workbook, mais dont la relation n’est pas une relation parent-enfant, ne sont pas des dépendants directs ou des antécédents directs.

Cet article fournit des exemples de code qui récupèrent des antécédents directs et des dépendances directes des formules à l’aide de l Excel API JavaScript. Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` [Range Object (interface API JavaScript pour Excel).](/javascript/api/excel/excel.range)

## <a name="get-the-direct-precedents-of-a-formula"></a>Obtenir les antécédents directs d’une formule

Recherchez les cellules précédentes directes d’une formule [avec Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--). `Range.getDirectPrecedents` renvoie un `WorkbookRangeAreas` objet. Cet objet contient les adresses de tous les précédents directs du manuel. Il possède un objet `RangeAreas` distinct pour chaque feuille de calcul contenant au moins un précédent de formule. Pour plus d’informations sur l’utilisation de l’objet, voir Work `RangeAreas` [with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).

La capture d’écran suivante montre le résultat de la sélection du bouton **Suivi des antécédents** dans Excel’interface utilisateur. Ce bouton dessine une flèche entre les cellules précédentes et la cellule sélectionnée. La cellule sélectionnée, **E3,** contient la formule « =C3 * D3 », c’est pourquoi **C3** et **D3** sont des cellules précédentes. Contrairement au bouton Excel’interface utilisateur, `getDirectPrecedents` la méthode ne dessine pas de flèches.

![Cellules précédentes de suivi des flèches dans Excel’interface utilisateur.](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> La `getDirectPrecedents` méthode ne peut pas récupérer les cellules précédentes dans les workbooks.

L’exemple de code suivant obtient les antécédents directs de la plage active, puis modifie la couleur d’arrière-plan de ces cellules précédentes en jaune.

```js
Excel.run(function (context) {
    // Precedents are cells that provide data to the selected formula.
    var range = context.workbook.getActiveCell();
    var directPrecedents = range.getDirectPrecedents();
    range.load("address");
    directPrecedents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct precedent cells of ${range.address}:`);

            // Use the direct precedents API to loop through precedents of the active cell.
            for (var i = 0; i < directPrecedents.areas.items.length; i++) {
              // Highlight and print out the address of each precedent cell.
              directPrecedents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directPrecedents.areas.items[i].address}`);
            }
        });
}).catch(errorHandlerFunction);
```

## <a name="get-the-direct-dependents-of-a-formula"></a>Obtenir les dépendants directs d’une formule

Recherchez les cellules dépendantes directes d’une formule [avec Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__). Like `Range.getDirectPrecedents` , renvoie également un `Range.getDirectDependents` `WorkbookRangeAreas` objet. Cet objet contient les adresses de tous les dépendants directs dans le manuel. Il possède un objet `RangeAreas` distinct pour chaque feuille de calcul contenant au moins une formule dépendante. Pour plus d’informations sur l’utilisation de l’objet, voir Work `RangeAreas` [with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).

La capture d’écran suivante montre le résultat de la sélection du bouton **Dépendants** du suivi dans Excel’interface utilisateur. Ce bouton dessine une flèche entre les cellules dépendantes et la cellule sélectionnée. La cellule sélectionnée, **D3,** a la cellule **E3** comme dépendant. **E3** contient la formule « =C3 * D3 ». Contrairement au bouton Excel’interface utilisateur, `getDirectDependents` la méthode ne dessine pas de flèches.

![Cellules dépendantes du suivi des flèches dans Excel’interface utilisateur.](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> La `getDirectDependents` méthode ne peut pas récupérer les cellules dépendantes dans les workbooks.

L’exemple de code suivant obtient les dépendants directs de la plage active, puis modifie la couleur d’arrière-plan de ces cellules dépendantes en jaune.

```js
Excel.run(function (context) {
    // Direct dependents are cells that contain formulas that refer to other cells.
    var range = context.workbook.getActiveCell();
    var directDependents = range.getDirectDependents();
    range.load("address");
    directDependents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct dependent cells of ${range.address}:`);
    
            // Use the direct dependents API to loop through direct dependents of the active cell.
            for (var i = 0; i < directDependents.areas.items.length; i++) {
              // Highlight and print the address of each dependent cell.
              directDependents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directDependents.areas.items[i].address}`);
            }
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide Excel API JavaScript](excel-add-ins-cells.md)
- [Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md)
