---
title: Supprimer les doublons à l’aide Excel API JavaScript
description: Découvrez comment utiliser l’API JavaScript Excel pour supprimer les doublons.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 03ef1c15f214fcbd6a81d0b4f557e9816175d3b7
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671132"
---
# <a name="remove-duplicates-using-the-excel-javascript-api"></a>Supprimer les doublons à l’aide Excel API JavaScript

Cet article fournit un exemple de code qui supprime les entrées en double dans une plage à l’aide de l Excel API JavaScript. Pour obtenir la liste complète des propriétés et méthodes que l’objet prend en `Range` charge, [voir Excel. Classe Range](/javascript/api/excel/excel.range).

## <a name="remove-rows-with-duplicate-entries"></a>Supprimer des lignes avec des entrées en double

La [méthode Range.removeDuplicates](/javascript/api/excel/excel.range#removeDuplicates_columns__includesHeader_) supprime les lignes avec des entrées en double dans les colonnes spécifiées. La méthode passe par chaque ligne de la plage, de l’index à la valeur la plus faible à l’index à valeur la plus élevée de la plage (du haut vers le bas). Une rangée est supprimée si une valeur dans sa/ses colonne(s) spécifiée(s) apparue(s) plus tôt dans la plage. Les rangées de la plage en-dessous de la rangée supprimée sont déplacées. `removeDuplicates` n’affecte pas la position des cellules en dehors de la rangée.

`removeDuplicates`prend un `number[]` représentant les indices de la colonne qui sont vérifiés pour les doublons. Ce tableau est à base zéro et lié à la rangée, et non à la feuille de calcul. La méthode prend également un paramètre booléen qui spécifie si la première ligne est un en-tête. Lorsque **true**, la rangée du dessus est ignorée lorsque les doublons sont pris en considération. La méthode renvoie un objet qui spécifie le nombre de lignes supprimées et le nombre de lignes `removeDuplicates` `RemoveDuplicatesResult` uniques restantes.

Lorsque vous utilisez la méthode `removeDuplicates` d’une plage, gardez ce qui suit à l’esprit.

- `removeDuplicates`considère les valeurs de cellule, et non les résultats de la fonction. Si deux fonctions différentes évaluent le même résultat, les valeurs de la cellule ne sont pas considérées comme doublons.
- Les cellules vides ne sont pas ignorées par`removeDuplicates`. La valeur d’une cellule vide est traitée comme toute autre valeur. Cela signifie que les rangées vides contenues au sein de la plage seront incluses dans le `RemoveDuplicatesResult`.

L’exemple de code suivant montre la suppression des entrées avec des valeurs en double dans la première colonne.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

### <a name="data-before-duplicate-entries-are-removed"></a>Données avant la suppression des entrées en double

![Données dans Excel méthode de suppression des doublons de la plage a été exécuté.](../images/excel-ranges-remove-duplicates-before.png)

### <a name="data-after-duplicate-entries-are-removed"></a>Données après suppression des entrées en double

![Données dans Excel une fois que la méthode des doublons de suppression de plage a été exécuté.](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide Excel API JavaScript](excel-add-ins-cells.md)
- [Couper, copier et coller des plages à l’aide de l Excel API JavaScript](excel-add-ins-ranges-cut-copy-paste.md)
- [Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md)
