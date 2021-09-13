---
title: Utiliser des cellules à l’aide Excel API JavaScript.
description: Découvrez la Excel de l’API JavaScript d’une cellule et découvrez comment utiliser des cellules.
ms.date: 04/16/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 74603727c5944583f55e77c75589f31ffbdffb21
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150451"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a>Utiliser des cellules à l’aide Excel API JavaScript

L’API JavaScript Excel ne comprend pas d’objet ou de classe « Cellule ». Au lieu de cela, Excel cellules sont `Range` des objets. Une cellule individuelle dans l’interface utilisateur d’Excel se traduit par un objet`Range` avec une cellule dans l’API JavaScript Excel.

Un `Range` objet peut également contenir plusieurs cellules contiguës. Les cellules contiguës forment un rectangle non abandonné (y compris des lignes ou des colonnes). Pour en savoir plus sur l’utilisation de cellules qui ne sont pas contiguës, voir Travailler avec des cellules non contiguës à l’aide de l’objet [RangeAreas](#work-with-discontiguous-cells-using-the-rangeareas-object).

Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` [Range Object (interface API JavaScript pour Excel).](/javascript/api/excel/excel.range)

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a>Utiliser des cellules peuigues à l’aide de l’objet RangeAreas

[L’objet RangeAreas permet](/javascript/api/excel/excel.rangeareas) à votre add-in d’effectuer des opérations sur plusieurs plages à la fois. Ces plages peuvent être contiguës, mais elles n’en ont pas besoin. `RangeAreas`sont abordés plus loin dans l’article[Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md).

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Obtenir une plage à l’aide de Excel API JavaScript](excel-add-ins-ranges-get.md)
- [Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md)
