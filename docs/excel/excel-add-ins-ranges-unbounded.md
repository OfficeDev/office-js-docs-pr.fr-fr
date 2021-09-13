---
title: Lire ou écrire dans une plage non limite à l’aide de l Excel API JavaScript
description: Découvrez comment utiliser l’API JavaScript Excel pour lire ou écrire dans une plage non limite.
ms.date: 04/05/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: a7b2a564377d0dab73d4f3ad6d3aacf2219ddeae
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152275"
---
# <a name="read-or-write-to-an-unbounded-range-using-the-excel-javascript-api"></a>Lire ou écrire dans une plage non limite à l’aide de l Excel API JavaScript

Cet article explique comment lire et écrire dans une plage non limite à l’Excel API JavaScript. Pour obtenir la liste complète des propriétés et méthodes que l’objet prend en `Range` charge, [voir Excel. Classe Range](/javascript/api/excel/excel.range).

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
- [Utiliser des cellules à l’aide Excel API JavaScript](excel-add-ins-cells.md)
- [Lire ou écrire dans une grande plage à l’aide de l Excel API JavaScript](excel-add-ins-ranges-large.md)
- [Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md)
