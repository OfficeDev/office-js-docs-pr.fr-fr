---
title: Plages de groupes à l’aide Excel API JavaScript
description: Découvrez comment grouper des lignes ou des colonnes d’une plage pour créer un plan à l’aide Excel API JavaScript.
ms.date: 04/05/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 9ec3f9e23f5099c703fbbf53fdc6fbb800acba6d
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149264"
---
# <a name="group-ranges-for-an-outline-using-the-excel-javascript-api"></a>Plages de groupe pour un plan à l’aide de l Excel API JavaScript

Cet article fournit un exemple de code qui montre comment grouper des plages pour un plan à l’aide de l Excel API JavaScript. Pour obtenir la liste complète des propriétés et méthodes que l’objet prend en `Range` charge, [voir Excel. Classe Range](/javascript/api/excel/excel.range).

## <a name="group-rows-or-columns-of-a-range-for-an-outline"></a>Grouper des lignes ou des colonnes d’une plage pour un plan

Les lignes ou colonnes d’une plage peuvent être regroupées pour créer un [plan.](https://support.microsoft.com/office/08ce98c4-0063-4d42-8ac7-8278c49e9aff) Ces groupes peuvent être réduire et développés pour masquer et afficher les cellules correspondantes. Cela facilite l’analyse rapide des données de première ligne. Utilisez [Range.group pour](/javascript/api/excel/excel.range#group_groupOption_) effectuer ces groupes de plan.

Un plan peut avoir une hiérarchie, où des groupes plus petits sont imbrmbrés sous des groupes plus grands. Cela permet d’afficher le plan à différents niveaux. La modification du niveau de plan visible peut être effectuée par programme via la [méthode Worksheet.showOutlineLevels.](/javascript/api/excel/excel.worksheet#showOutlineLevels_rowLevels__columnLevels_) Notez que Excel ne prend en charge que huit niveaux de groupes de plan.

L’exemple de code suivant crée un plan avec deux niveaux de groupes pour les lignes et les colonnes. L’image suivante montre les regroupements de ce plan. Dans l’exemple de code, les plages regroupées n’incluent pas la ligne ou la colonne du contrôle de plan (les « Totaux » pour cet exemple). Un groupe définit ce qui sera réduire, et non la ligne ou la colonne avec le contrôle.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Group the larger, main level. Note that the outline controls
    // will be on row 10, meaning 4-9 will collapse and expand.
    sheet.getRange("4:9").group(Excel.GroupOption.byRows);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on rows 6 and 9, meaning 4-5 and 7-8 will collapse and expand.
    sheet.getRange("4:5").group(Excel.GroupOption.byRows);
    sheet.getRange("7:8").group(Excel.GroupOption.byRows);

    // Group the larger, main level. Note that the outline controls
    // will be on column R, meaning C-Q will collapse and expand.
    sheet.getRange("C:Q").group(Excel.GroupOption.byColumns);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on columns G, L, and R, meaning C-F, H-K, and M-P will collapse and expand.
    sheet.getRange("C:F").group(Excel.GroupOption.byColumns);
    sheet.getRange("H:K").group(Excel.GroupOption.byColumns);
    sheet.getRange("M:P").group(Excel.GroupOption.byColumns);
    return context.sync();
}).catch(errorHandlerFunction);
```

![Plage avec un plan à deux niveaux à deux dimensions.](../images/excel-outline.png)

## <a name="remove-grouping-from-rows-or-columns-of-a-range"></a>Supprimer le regroupement des lignes ou des colonnes d’une plage

Pour regrouper un groupe de lignes ou de colonnes, utilisez [la méthode Range.ungroup.](/javascript/api/excel/excel.range#ungroup_groupOption_) Cela supprime le niveau le plus à l’extérieur du plan. Si plusieurs groupes du même type de ligne ou de colonne sont au même niveau dans la plage spécifiée, tous ces groupes sont désgroupés.

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide Excel API JavaScript](excel-add-ins-cells.md)
- [Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md)
