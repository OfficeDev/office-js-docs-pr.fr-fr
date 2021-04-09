---
title: Utiliser des antécédents de formule à l’aide de l’API JavaScript pour Excel
description: Découvrez comment utiliser l’API JavaScript excel pour récupérer les antécédents de formule.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0d21ae411615a22873a0f4dda185984f6191ac8e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652840"
---
# <a name="get-formula-precedents-using-the-excel-javascript-api"></a>Obtenir des antécédents de formule à l’aide de l’API JavaScript pour Excel

Cet article fournit un exemple de code qui récupère les antécédents de formule à l’aide de l’API JavaScript pour Excel. Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` la classe [Excel.Range.](/javascript/api/excel/excel.range)

## <a name="get-formula-precedents"></a>Obtenir des antécédents de formule

Une formule Excel fait souvent référence à d’autres cellules. Lorsqu’une cellule fournit des données à une formule, elle est appelée formule « antécédent ». Pour en savoir plus sur les fonctionnalités Excel liées aux relations entre les cellules, voir Afficher les [relations entre les formules et les cellules.](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507) 

Avec [Range.getDirectPrecedents,](/javascript/api/excel/excel.range#getdirectprecedents--)votre add-in peut localiser les cellules précédentes directes d’une formule. `Range.getDirectPrecedents` renvoie un `WorkbookRangeAreas` objet. Cet objet contient les adresses de tous les antécédents dans le manuel. Il possède un objet `RangeAreas` distinct pour chaque feuille de calcul contenant au moins un précédent de formule. Pour plus d’informations sur l’utilisation de l’objet, voir Work `RangeAreas` [with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).

Dans l’interface utilisateur Excel, le bouton **Suivi des antécédents** dessine une flèche entre les cellules précédentes et la formule sélectionnée. Contrairement au bouton de l’interface utilisateur Excel, `getDirectPrecedents` la méthode ne dessine pas de flèches. 

> [!IMPORTANT]
> La `getDirectPrecedents` méthode ne peut pas récupérer les cellules précédentes dans les workbooks. 

L’exemple de code suivant obtient les antécédents directs de la plage active, puis modifie la couleur d’arrière-plan de ces cellules précédentes en jaune. 

> [!NOTE]
> La plage active doit contenir une formule qui fait référence à d’autres cellules du même workbook pour que la mise en surbrillance fonctionne correctement. 

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
        })
        .then(context.sync);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide de l’API JavaScript pour Excel](excel-add-ins-cells.md)
- [Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md)
