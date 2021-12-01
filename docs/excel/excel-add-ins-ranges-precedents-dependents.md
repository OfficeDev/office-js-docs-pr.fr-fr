---
title: Utiliser les antécédents et les dépendances de formule à l’aide Excel API JavaScript
description: Découvrez comment utiliser l’API JavaScript Excel pour récupérer les antécédents et les dépendances de formule.
ms.date: 11/30/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 60da910879fc48f1564d43cf3f87c2a5bf930fbe
ms.sourcegitcommit: 5daf91eb3be99c88b250348186189f4dc1270956
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/01/2021
ms.locfileid: "61242060"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a>Obtenir des antécédents et des dépendances de formule à l’aide Excel API JavaScript

Excel formules font souvent référence à d’autres cellules. Ces références entre cellules sont appelées « antécédents » et « dépendants ». Un précédent est une cellule qui fournit des données à une formule. Une cellule dépendante est une cellule qui contient une formule qui fait référence à d’autres cellules. Pour en savoir plus sur Excel fonctionnalités liées aux relations entre les cellules, voir Afficher les relations entre les [formules et les cellules.](https://support.microsoft.com/office/a59bef2b-3701-46bf-8ff1-d3518771d507)

Une cellule précédente peut avoir ses propres cellules précédentes. Chaque cellule précédente de cette chaîne de précédents est toujours un antécédent de la cellule d’origine. La même relation existe pour les dépendants. Toute cellule affectée par une autre cellule dépend de cette cellule. Un « précédent direct » est le premier groupe de cellules précédent dans cette séquence, similaire au concept de parents dans une relation parent-enfant. Un « dépendant direct » est le premier groupe dépendant de cellules dans une séquence, semblable aux enfants d’une relation parent-enfant.

Cet article fournit des exemples de code qui récupèrent les antécédents et les dépendances des formules à l’aide Excel API JavaScript. Pour obtenir la liste complète des propriétés et méthodes que l’objet prend en charge, voir `Range` [Range Object (interface API JavaScript pour Excel).](/javascript/api/excel/excel.range)

## <a name="get-the-precedents-of-a-formula"></a>Obtenir les antécédents d’une formule

Recherchez les cellules précédentes d’une formule [avec Range.getPrecedents](/javascript/api/excel/excel.range#getPrecedents__). `Range.getPrecedents` renvoie un `WorkbookRangeAreas` objet. Cet objet contient les adresses de tous les précédents dans le manuel. Il possède un objet `RangeAreas` distinct pour chaque feuille de calcul contenant au moins un précédent de formule. Pour en savoir plus sur l’objet, voir Travailler avec plusieurs plages simultanément `RangeAreas` dans Excel des [modules.](excel-add-ins-multiple-ranges.md)

Pour localiser uniquement les cellules précédentes directes d’une formule, [utilisez Range.getDirectPrecedents](/javascript/api/excel/excel.range#getDirectPrecedents__). `Range.getDirectPrecedents` fonctionne comme `Range.getPrecedents` et renvoie un objet contenant les `WorkbookRangeAreas` adresses de précédents directs.

La capture d’écran suivante montre le résultat de la sélection du bouton Suivi **des antécédents** dans Excel’interface utilisateur. Ce bouton dessine une flèche entre les cellules précédentes et la cellule sélectionnée. La cellule sélectionnée, **E3,** contient la formule « =C3 * D3 », c’est-à-dire que **C3** et **D3** sont des cellules précédentes. Contrairement au Excel’interface utilisateur, les méthodes et les méthodes `getPrecedents` `getDirectPrecedents` ne dessinent pas de flèches.

![Cellules précédentes de suivi des flèches dans Excel’interface utilisateur.](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> Les `getPrecedents` méthodes et les méthodes ne `getDirectPrecedents` récupèrent pas les cellules précédentes dans les workbooks.

L’exemple de code suivant montre comment travailler avec les `Range.getPrecedents` méthodes `Range.getDirectPrecedents` et les méthodes. L’exemple obtient les antécédents de la plage active, puis modifie la couleur d’arrière-plan de ces cellules précédentes. La couleur d’arrière-plan des cellules précédentes directes est jaune et la couleur d’arrière-plan des autres cellules précédentes est définie sur orange.

```js
// This code sample shows how to find and highlight the precedents 
// and direct precedents of the currently selected cell.
Excel.run(function (context) {
  var range = context.workbook.getActiveCell();
  // Precedents are all cells that provide data to the selected formula.
  var precedents = range.getPrecedents();
  // Direct precedents are the parent cells, or the first preceding group of cells that provide data to the selected formula.    
  var directPrecedents = range.getDirectPrecedents();

  range.load("address");
  precedents.areas.load("address");
  directPrecedents.areas.load("address");
  
  return context.sync()
    .then(function () {
      console.log(`All precedent cells of ${range.address}:`);
      
      // Use the precedents API to loop through all precedents of the active cell.
      for (var i = 0; i < precedents.areas.items.length; i++) {
        // Highlight and print out the address of all precedent cells.
        precedents.areas.items[i].format.fill.color = "Orange";
        console.log(`  ${precedents.areas.items[i].address}`);
      }

      console.log(`Direct precedent cells of ${range.address}:`);

      // Use the direct precedents API to loop through direct precedents of the active cell.
      for (var i = 0; i < directPrecedents.areas.items.length; i++) {
        // Highlight and print out the address of each direct precedent cell.
        directPrecedents.areas.items[i].format.fill.color = "Yellow";
        console.log(`  ${directPrecedents.areas.items[i].address}`);
      }
    });
}).catch(errorHandlerFunction);
```

## <a name="get-the-direct-dependents-of-a-formula"></a>Obtenir les dépendants directs d’une formule

Recherchez les cellules dépendantes directes d’une formule [avec Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__). Like `Range.getDirectPrecedents` , renvoie également un `Range.getDirectDependents` `WorkbookRangeAreas` objet. Cet objet contient les adresses de tous les dépendants directs du manuel. Il possède un objet `RangeAreas` distinct pour chaque feuille de calcul contenant au moins une formule dépendante. Pour plus d’informations sur l’utilisation de l’objet, voir Work `RangeAreas` [with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).

La capture d’écran suivante montre le résultat de la sélection du bouton **Dépendants** du suivi dans Excel’interface utilisateur. Ce bouton dessine une flèche entre les cellules dépendantes et la cellule sélectionnée. La cellule sélectionnée, **D3,** a la cellule **E3** comme dépendant. **E3** contient la formule « =C3 * D3 ». Contrairement au bouton Excel’interface utilisateur, `getDirectDependents` la méthode ne dessine pas de flèches.

![Cellules dépendantes du suivi des flèches dans Excel’interface utilisateur.](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> La `getDirectDependents` méthode ne récupère pas les cellules dépendantes dans les workbooks.

L’exemple de code suivant obtient les dépendants directs de la plage active, puis modifie la couleur d’arrière-plan de ces cellules dépendantes en jaune.

```js
// This code sample shows how to find and highlight the dependents of the currently selected cell.
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
- [Utiliser des cellules à l’aide de Excel API JavaScript](excel-add-ins-cells.md)
- [Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md)
