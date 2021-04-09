---
title: Rechercher une chaîne à l’aide de l’API JavaScript pour Excel
description: Découvrez comment trouver une chaîne dans une plage à l’aide de l’API JavaScript pour Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9b649bb249cd24d7578bc4f8285e5d0a23d0e4cd
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652821"
---
# <a name="find-a-string-within-a-range-using-the-excel-javascript-api"></a>Rechercher une chaîne dans une plage à l’aide de l’API JavaScript pour Excel

Cet article fournit un exemple de code qui trouve une chaîne dans une plage à l’aide de l’API JavaScript pour Excel. Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` la classe [Excel.Range.](/javascript/api/excel/excel.range)

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="match-a-string-within-a-range"></a>Faire correspondre une chaîne dans une plage

L’objet `Range` dispose d’une méthode`find` pour rechercher une chaîne spécifiée dans la plage. Elle renvoie la plage de la première cellule avec le texte correspondant.

L’exemple de code suivant trouve la première cellule contenant une valeur égale à la chaîne **Nourriture** et connecte son adresse à la console. Notez que `find` génère une erreur `ItemNotFound` si la chaîne spécifiée n’existe pas dans la plage. Si vous pensez que la chaîne spécifiée peut ne pas exister dans la plage, utilisez la méthode[findAllOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) à la place, pour que votre code gère ce scénario plus facilement.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var table = sheet.tables.getItem("ExpensesTable");
    var searchRange = table.getRange();
    var foundRange = searchRange.find("Food", {
        completeMatch: true, // find will match the whole cell value
        matchCase: false, // find will not match case
        searchDirection: Excel.SearchDirection.forward // find will start searching at the beginning of the range
    });

    foundRange.load("address");
    return context.sync()
        .then(function() {
            console.log(foundRange.address);
    });
}).catch(errorHandlerFunction);
```

Lorsque la méthode `find` est appelée sur une plage représentant une cellule simple, la feuille de calcul entière est recherchée. La recherche commence à cette cellule et continue dans la direction spécifiée par `SearchCriteria.searchDirection`, revenant à la ligne à la fin de la feuille de calcul si nécessaire.

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Utiliser des cellules à l’aide de l’API JavaScript pour Excel](excel-add-ins-cells.md)
- [Rechercher des cellules spéciales dans une plage à l’aide de l’API JavaScript pour Excel](excel-add-ins-ranges-special-cells.md)
