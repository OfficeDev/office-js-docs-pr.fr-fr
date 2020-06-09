---
title: Appeler des API Microsoft Excel à partir d’une fonction personnalisée
description: Découvrez les API Microsoft Excel que vous pouvez appeler à partir de votre fonction personnalisée.
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: a25d3f151f648560ee24a3da3f689cb9767bd52a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609803"
---
# <a name="call-microsoft-excel-apis-from-a-custom-function"></a>Appeler des API Microsoft Excel à partir d’une fonction personnalisée

Appelez les API Excel Office. js à partir de vos fonctions personnalisées pour obtenir des données de plage et obtenir davantage de contexte pour vos calculs.

L’appel des API Office. js via une fonction personnalisée peut être utile dans les cas suivants :

- Une fonction personnalisée doit obtenir des informations à partir d’Excel avant le calcul. Ces informations peuvent inclure des propriétés de document, des formats de plage, des parties XML personnalisées, un nom de classeur ou d’autres informations spécifiques à Excel.
- Une fonction personnalisée définit le format numérique de la cellule pour les valeurs renvoyées après le calcul.

## <a name="code-sample"></a>Exemple de code

Pour appeler les API Office. js, vous avez d’abord besoin d’un contexte. Utilisez l' `Excel.RequestContext` objet pour obtenir un contexte. Ensuite, utilisez le contexte pour appeler les API dont vous avez besoin dans le classeur.

L’exemple de code suivant montre comment obtenir une plage de valeurs du classeur.

```JavaScript
/**
 * @customfunction
 * @param address range's address
 **/
async function getRangeValue (address) {
 var context = new Excel.RequestContext();
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load();
 await context.sync();
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-officejs-through-a-custom-function"></a>Limitations de l’appel d’Office. js via une fonction personnalisée

N’appelez pas les API Office. js à partir d’une fonction personnalisée qui modifie l’environnement d’Excel. Cela signifie que vos fonctions personnalisées ne doivent pas effectuer les opérations suivantes :

- Insérer, supprimer ou mettre en forme des cellules dans la feuille de calcul.
- Modifier la valeur d’une autre cellule.
- Déplacer, renommer, supprimer ou ajouter des feuilles dans un classeur.
- Modifier les options d’environnement, telles que le mode de calcul ou les affichages d’écran.
- Ajouter des noms à un classeur.
- Définir des propriétés ou exécuter la plupart des méthodes.

La modification d’Excel peut entraîner une dégradation des performances, des délais et des boucles infinies. Les calculs de fonctions personnalisées ne doivent pas s’exécuter lorsqu’un recalcul Excel a lieu, car cela entraînera des résultats imprévisibles.

Au lieu de cela, modifiez Excel à partir du contexte d’un bouton de ruban ou d’un volet de tâches.

## <a name="next-steps"></a>Étapes suivantes

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>Voir aussi

- [Partager des données et des événements entre des fonctions personnalisées Excel et un didacticiel de volet de tâches](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
