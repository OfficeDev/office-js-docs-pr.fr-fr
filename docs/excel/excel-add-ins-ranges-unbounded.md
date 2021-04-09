---
title: Lire ou écrire dans une plage non limite à l’aide de l’API JavaScript pour Excel
description: Découvrez comment utiliser l’API JavaScript pour Excel pour lire ou écrire dans une plage non limite.
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f7be2efc3e069ea3451088608ca5255a632ef863
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652820"
---
# <a name="read-or-write-to-an-unbounded-range-using-the-excel-javascript-api"></a>Lire ou écrire dans une plage non limite à l’aide de l’API JavaScript pour Excel

Cet article explique comment lire et écrire dans une plage non limite avec l’API JavaScript pour Excel. Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` la classe [Excel.Range.](/javascript/api/excel/excel.range)

Une adresse de plage non limite est une adresse de plage qui spécifie des colonnes entières ou des lignes entières. Par exemple :

- Adresses de plage composées de colonnes entières :<ul><li>`C:C`</li><li>`A:F`</li></ul>
- Adresses de plage composées de lignes entières :<ul><li>`2:2`</li><li>`1:4`</li></ul>

## <a name="read-an-unbounded-range"></a>Lire une plage non liée

Lorsque l’API effectue une demande de récupération d’une plage non liée (par exemple, `getRange('C:C')`), la réponse contient des valeurs `null` pour les propriétés définies au niveau des cellules, telles que `values`, `text`, `numberFormat` et `formula`. Les autres propriétés de la plage, telles que `address` et `cellCount`, contiennent des valeurs valides pour la plage non liée.

## <a name="write-to-an-unbounded-range"></a>Écrire dans une plage non liée

Vous ne pouvez pas définir de propriétés au niveau de la cellule telles que , et sur une plage non limite, car la demande d’entrée `values` `numberFormat` est trop `formula` grande. Par exemple, l’exemple de code suivant n’est pas valide, car il tente de spécifier une plage `values` non limite. L’API renvoie une erreur si vous tentez de définir des propriétés au niveau de la cellule pour une plage non limite.

```js
// Note: This code sample attempts to specify `values` for an unbounded range, which is not a valid request. The sample will return an error. 
var range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide de l’API JavaScript pour Excel](excel-add-ins-cells.md)
- [Lire ou écrire dans une grande plage à l’aide de l’API JavaScript pour Excel](excel-add-ins-ranges-large.md)
- [Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md)
