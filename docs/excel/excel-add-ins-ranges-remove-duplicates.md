---
title: Supprimer les doublons à l’aide de l’API JavaScript Excel
description: Découvrez comment utiliser l’API JavaScript Excel pour supprimer les doublons.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 9ece7c9f35b341dbb8d0d90e8ca4bda5215580ed
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889141"
---
# <a name="remove-duplicates-using-the-excel-javascript-api"></a>Supprimer les doublons à l’aide de l’API JavaScript Excel

Cet article fournit un exemple de code qui supprime les entrées en double dans une plage à l’aide de l’API JavaScript Excel. Pour obtenir la liste complète des propriétés et méthodes prises en charge par l’objet `Range` , consultez la [classe Excel.Range](/javascript/api/excel/excel.range).

## <a name="remove-rows-with-duplicate-entries"></a>Supprimer des lignes avec des entrées en double

La méthode [Range.removeDuplicates](/javascript/api/excel/excel.range#excel-excel-range-removeduplicates-member(1)) supprime les lignes avec des entrées en double dans les colonnes spécifiées. La méthode passe par chaque ligne de la plage, de l’index à valeur inférieure à l’index à valeur la plus élevée de la plage (de haut en bas). Une rangée est supprimée si une valeur dans sa/ses colonne(s) spécifiée(s) apparue(s) plus tôt dans la plage. Les rangées de la plage en-dessous de la rangée supprimée sont déplacées. `removeDuplicates` n’affecte pas la position des cellules en dehors de la rangée.

`removeDuplicates`prend un `number[]` représentant les indices de la colonne qui sont vérifiés pour les doublons. Ce tableau est à base zéro et lié à la rangée, et non à la feuille de calcul. La méthode accepte également un paramètre booléen qui spécifie si la première ligne est un en-tête. Quand `true`, la ligne supérieure est ignorée lors de l’examen des doublons. La `removeDuplicates` méthode renvoie un `RemoveDuplicatesResult` objet qui spécifie le nombre de lignes supprimées et le nombre de lignes uniques restantes.

Lorsque vous utilisez la méthode d’une `removeDuplicates` plage, gardez à l’esprit ce qui suit.

- `removeDuplicates`considère les valeurs de cellule, et non les résultats de la fonction. Si deux fonctions différentes évaluent le même résultat, les valeurs de la cellule ne sont pas considérées comme doublons.
- Les cellules vides ne sont pas ignorées par`removeDuplicates`. La valeur d’une cellule vide est traitée comme toute autre valeur. Cela signifie que les rangées vides contenues au sein de la plage seront incluses dans le `RemoveDuplicatesResult`.

L’exemple de code suivant montre la suppression d’entrées avec des valeurs en double dans la première colonne.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("B2:D11");

    let deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    await context.sync();

    console.log(deleteResult.removed + " entries with duplicate names removed.");
    console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
});
```

### <a name="data-before-duplicate-entries-are-removed"></a>Données avant la suppression des entrées en double

![Les données dans Excel avant l’exécution de la méthode supprimer les doublons de la plage ont été exécutées.](../images/excel-ranges-remove-duplicates-before.png)

### <a name="data-after-duplicate-entries-are-removed"></a>Données après la suppression des entrées dupliquées

![Données dans Excel après l’exécution de la méthode supprimer les doublons de la plage.](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide de l’API JavaScript Excel](excel-add-ins-cells.md)
- [Couper, copier et coller des plages à l’aide de l’API JavaScript Excel](excel-add-ins-ranges-cut-copy-paste.md)
- [Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md)
