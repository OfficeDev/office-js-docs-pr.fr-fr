---
title: Utiliser des modèles de formule précédents et dépendants à l’aide de l’API JavaScript Excel
description: Découvrez comment utiliser l’API JavaScript Excel pour récupérer les antécédents de formule et les dépendances.
ms.date: 05/19/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: ca432b7eb6825781960e995af2ed2193c7caa5e2
ms.sourcegitcommit: 4ca3334f3cefa34e6b391eb92a429a308229fe89
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2022
ms.locfileid: "65628095"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a>Obtenir les antécédents et les dépendances de formule à l’aide de l’API JavaScript Excel

Excel formules font souvent référence à d’autres cellules. Ces références entre cellules sont appelées « précédents » et « dépendants ». Un précédent est une cellule qui fournit des données à une formule. Une cellule dépendante est une cellule qui contient une formule qui fait référence à d’autres cellules. Pour en savoir plus sur Excel fonctionnalités liées aux relations entre les cellules, consultez [Afficher les relations entre les formules et les cellules](https://support.microsoft.com/office/a59bef2b-3701-46bf-8ff1-d3518771d507).

Une cellule précédente peut avoir ses propres cellules de précédent. Chaque cellule précédente de cette chaîne de précédents est toujours un précédent de la cellule d’origine. La même relation existe pour les personnes dépendantes. Toute cellule affectée par une autre cellule dépend de cette cellule. Un « précédent direct » est le premier groupe de cellules précédent dans cette séquence, semblable au concept de parents dans une relation parent-enfant. Un « dépendant direct » est le premier groupe dépendant de cellules d’une séquence, semblable aux enfants d’une relation parent-enfant.

Cet article fournit des exemples de code qui récupèrent des précédents et des dépendances de formules à l’aide de l’API JavaScript Excel. Pour obtenir la liste complète des propriétés et méthodes prises en charge par l’objet`Range`, consultez [Range Object (Interface API JavaScript pour Excel).](/javascript/api/excel/excel.range)

## <a name="get-the-precedents-of-a-formula"></a>Obtenir les précédents d’une formule

Recherchez les cellules précédentes d’une formule avec [Range.getPrecedents](/javascript/api/excel/excel.range#excel-excel-range-getprecedents-member(1)). `Range.getPrecedents` renvoie un `WorkbookRangeAreas` objet. Cet objet contient les adresses de tous les précédents du classeur. Il a un objet distinct `RangeAreas` pour chaque feuille de calcul contenant au moins un précédent de formule. Pour en savoir plus sur l’objet`RangeAreas`, consultez [Utiliser plusieurs plages simultanément dans Excel compléments](excel-add-ins-multiple-ranges.md).

Pour localiser uniquement les cellules précédentes directes d’une formule, utilisez [Range.getDirectPrecedents](/javascript/api/excel/excel.range#excel-excel-range-getdirectprecedents-member(1)). `Range.getDirectPrecedents` fonctionne comme `Range.getPrecedents` un `WorkbookRangeAreas` objet contenant les adresses des précédents directs.

La capture d’écran suivante montre le résultat de la sélection du bouton **Précédents** de trace dans l’interface utilisateur Excel. Ce bouton dessine une flèche à partir des cellules précédentes vers la cellule sélectionnée. La cellule sélectionnée, **E3**, contient la formule « =C3 * D3 », de sorte que **C3** et **D3** sont des cellules de précédent. Contrairement au bouton Excel’interface utilisateur, les méthodes et `getDirectPrecedents` les `getPrecedents` méthodes ne dessinent pas de flèches.

![Cellules précédentes de suivi de flèche dans l’interface utilisateur Excel.](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> Les `getPrecedents` méthodes et `getDirectPrecedents` les méthodes ne récupèrent pas les cellules précédentes dans les classeurs.

L’exemple de code suivant montre comment utiliser les méthodes et `Range.getDirectPrecedents` les `Range.getPrecedents` méthodes. L’exemple obtient les précédents pour la plage active, puis modifie la couleur d’arrière-plan de ces cellules précédentes. La couleur d’arrière-plan des cellules précédentes directes est définie sur jaune et la couleur d’arrière-plan des autres cellules précédentes est orange.

```js
// This code sample shows how to find and highlight the precedents 
// and direct precedents of the currently selected cell.
await Excel.run(async (context) => {
  let range = context.workbook.getActiveCell();
  // Precedents are all cells that provide data to the selected formula.
  let precedents = range.getPrecedents();
  // Direct precedents are the parent cells, or the first preceding group of cells that provide data to the selected formula.    
  let directPrecedents = range.getDirectPrecedents();

  range.load("address");
  precedents.areas.load("address");
  directPrecedents.areas.load("address");
  
  await context.sync();

  console.log(`All precedent cells of ${range.address}:`);
  
  // Use the precedents API to loop through all precedents of the active cell.
  for (let i = 0; i < precedents.areas.items.length; i++) {
    // Highlight and print out the address of all precedent cells.
    precedents.areas.items[i].format.fill.color = "Orange";
    console.log(`  ${precedents.areas.items[i].address}`);
  }

  console.log(`Direct precedent cells of ${range.address}:`);

  // Use the direct precedents API to loop through direct precedents of the active cell.
  for (let i = 0; i < directPrecedents.areas.items.length; i++) {
    // Highlight and print out the address of each direct precedent cell.
    directPrecedents.areas.items[i].format.fill.color = "Yellow";
    console.log(`  ${directPrecedents.areas.items[i].address}`);
  }
});
```

## <a name="get-the-dependents-of-a-formula"></a>Obtenir les dépendances d’une formule

Recherchez les cellules dépendantes d’une formule avec [Range.getDependents](/javascript/api/excel/excel.range#excel-excel-range-getdependents-member(1)). Like `Range.getPrecedents`, `Range.getDependents` renvoie également un `WorkbookRangeAreas` objet. Cet objet contient les adresses de toutes les personnes dépendantes du classeur. Il a un objet distinct `RangeAreas` pour chaque feuille de calcul contenant au moins une formule dépendante. Pour plus d’informations sur l’utilisation de l’objet`RangeAreas`, consultez [Travailler avec plusieurs plages simultanément dans Excel compléments](excel-add-ins-multiple-ranges.md).

Pour localiser uniquement les cellules dépendantes directes d’une formule, utilisez [Range.getDirectDependents](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1)). `Range.getDirectDependents` fonctionne comme `Range.getDependents` et retourne un `WorkbookRangeAreas` objet contenant les adresses des dépendants directs.

La capture d’écran suivante montre le résultat de la sélection du bouton **Dépendants de la trace** dans l’interface utilisateur Excel. Ce bouton dessine une flèche de la cellule sélectionnée vers les cellules dépendantes. La cellule sélectionnée, **D3**, a la cellule **E3** en tant que cellule dépendante. **E3** contient la formule « =C3 * D3 ». Contrairement au bouton Excel’interface utilisateur, les méthodes et `getDirectDependents` les `getDependents` méthodes ne dessinent pas de flèches.

![Flèche traçant les cellules dépendantes dans l’interface utilisateur Excel.](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> Les `getDependents` méthodes et `getDirectDependents` les méthodes ne récupèrent pas les cellules dépendantes dans les classeurs.

L’exemple de code suivant obtient les dépendants directs de la plage active, puis modifie la couleur d’arrière-plan de ces cellules dépendantes en jaune.

L’exemple de code suivant montre comment utiliser les méthodes et `Range.getDirectDependents` les `Range.getDependents` méthodes. L’exemple obtient les dépendants de la plage active, puis modifie la couleur d’arrière-plan de ces cellules dépendantes. La couleur d’arrière-plan des cellules dépendantes directes est définie sur jaune et la couleur d’arrière-plan des autres cellules dépendantes est orange.

```js
// This code sample shows how to find and highlight the dependents 
// and direct dependents of the currently selected cell.
await Excel.run(async (context) => {
    let range = context.workbook.getActiveCell();
    // Dependents are all cells that contain formulas that refer to other cells.
    let dependents = range.getDependents();  
    // Direct dependents are the child cells, or the first succeeding group of cells in a sequence of cells that refer to other cells.
    let directDependents = range.getDirectDependents();

    range.load("address");
    dependents.areas.load("address");    
    directDependents.areas.load("address");
    
    await context.sync();

    console.log(`All dependent cells of ${range.address}:`);
    
    // Use the dependents API to loop through all dependents of the active cell.
    for (let i = 0; i < dependents.areas.items.length; i++) {
      // Highlight and print out the addresses of all dependent cells.
      dependents.areas.items[i].format.fill.color = "Orange";
      console.log(`  ${dependents.areas.items[i].address}`);
    }

    console.log(`Direct dependent cells of ${range.address}:`);

    // Use the direct dependents API to loop through direct dependents of the active cell.
    for (let i = 0; i < directDependents.areas.items.length; i++) {
      // Highlight and print the address of each dependent cell.
      directDependents.areas.items[i].format.fill.color = "Yellow";
      console.log(`  ${directDependents.areas.items[i].address}`);
    }
});
```

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide de l’API JavaScript Excel](excel-add-ins-cells.md)
- [Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md)
