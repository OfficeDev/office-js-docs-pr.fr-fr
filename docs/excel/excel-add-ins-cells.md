---
title: Utiliser des cellules à l’aide de l’API JavaScript pour Excel.
description: Découvrez la définition de l’API JavaScript pour Excel d’une cellule et découvrez comment utiliser des cellules.
ms.date: 04/07/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 5fcfeeef52f17c22d13ed3c1a10851f1d8e69204
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652873"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a>Utiliser des cellules à l’aide de l’API JavaScript pour Excel

L’API JavaScript pour Excel n’a pas d’objet ou de classe « Cell ». Au lieu de cela, toutes les cellules Excel sont `Range` des objets. Une cellule individuelle dans l’interface utilisateur d’Excel se traduit par un objet avec une cellule dans `Range` l’API JavaScript pour Excel.

Un `Range` objet peut également contenir plusieurs cellules contiguës. Les cellules contiguës forment un rectangle non abandonné (y compris des lignes ou des colonnes). Pour en savoir plus sur l’utilisation de cellules qui ne sont pas contiguës, voir Travailler avec des cellules non contiguës à l’aide de l’objet [RangeAreas](#work-with-discontiguous-cells-using-the-rangeareas-object).

Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` la classe [Excel.Range.](/javascript/api/excel/excel.range)

## <a name="excel-javascript-apis-that-mention-cells"></a>API JavaScript Excel mentionnant les cellules

Même si l’API JavaScript pour Excel n’a pas d’objet ou de classe « Cell », un certain nombre de noms d’API mentionnent des cellules. Ces API contrôlent les propriétés des cellules telles que la couleur, la mise en forme du texte et la police.

La liste suivante des API JavaScript pour Excel fait référence à des cellules.

- [CellBorder](/javascript/api/excel/excel.cellborder)
- [CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)
- [CellProperties](/javascript/api/excel/excel.cellproperties)
- [CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)
- [CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)
- [CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)
- [CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)
- [CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)
- [ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)
- [SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a>Utiliser des cellules peuigues à l’aide de l’objet RangeAreas

[L’objet RangeAreas permet](/javascript/api/excel/excel.rangeareas) à votre add-in d’effectuer des opérations sur plusieurs plages à la fois. Ces plages peuvent être contiguës, mais elles n’en ont pas besoin. `RangeAreas`sont abordés plus loin dans l’article[Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md).

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Obtenir une plage à l’aide de l’API JavaScript pour Excel](excel-add-ins-ranges-get.md)
- [Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md)
