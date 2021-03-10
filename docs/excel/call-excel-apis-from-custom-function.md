---
title: Appeler des API JavaScript Excel à partir d’une fonction personnalisée
description: Découvrez les API JavaScript Excel que vous pouvez appeler à partir de votre fonction personnalisée.
ms.date: 03/05/2021
localization_priority: Normal
ms.openlocfilehash: 4be1b1ee8ea4ae8b2f5d1d27195be18f7aa841da
ms.sourcegitcommit: d153f6d4c3e01d63ed24aa1349be16fa8ad51218
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/10/2021
ms.locfileid: "50613905"
---
# <a name="call-excel-javascript-apis-from-a-custom-function"></a>Appeler des API JavaScript Excel à partir d’une fonction personnalisée

Appelez des API JavaScript Excel à partir de vos fonctions personnalisées pour obtenir des données de plage et obtenir plus de contexte pour vos calculs. L’appel d’API JavaScript Pour Excel via une fonction personnalisée peut être utile dans les cas de :

- Une fonction personnalisée doit obtenir des informations d’Excel avant le calcul. Ces informations peuvent inclure des propriétés de document, des formats de plage, des parties XML personnalisées, un nom de workbook ou d’autres informations spécifiques à Excel.
- Une fonction personnalisée définira le format numérique de la cellule pour les valeurs de retour après le calcul.

> [!IMPORTANT]
> Pour appeler des API JavaScript Excel à partir de votre fonction personnalisée, vous devez utiliser un runtime JavaScript partagé. Pour plus d’information, consultez [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="code-sample"></a>Exemple de code

Pour appeler des API JavaScript Excel à partir d’une fonction personnalisée, vous avez d’abord besoin d’un contexte. Utilisez [l’objet Excel.RequestContext](/javascript/api/excel/excel.requestcontext) pour obtenir un contexte. Utilisez ensuite le contexte pour appeler les API dont vous avez besoin dans le workbook.

L’exemple de code suivant montre comment utiliser pour obtenir une valeur à partir `Excel.RequestContext` d’une cellule dans le workbook. Dans cet exemple, le paramètre est transmis dans la méthode `address` [Worksheet.getRange](/javascript/api/excel/excel.worksheet#getRange_address_) de l’API JavaScript pour Excel et doit être entré sous forme de chaîne. Par exemple, la fonction personnalisée entrée dans l’interface utilisateur Excel doit suivre le modèle , où est l’adresse de la cellule à partir de laquelle récupérer `=CONTOSO.GETRANGEVALUE("A1")` `"A1"` la valeur.

```JavaScript
/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
async function getRangeValue(address) {
 // Retrieve the context object. 
 var context = new Excel.RequestContext();
 
 // Use the context object to access the cell at the input address. 
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load();
 await context.sync();
 
 // Return the value of the cell at the input address.
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a>Limitations de l’appel d’API JavaScript pour Excel via une fonction personnalisée

N’appelez pas les API JavaScript pour Excel à partir d’une fonction personnalisée qui modifie l’environnement d’Excel. Cela signifie que vos fonctions personnalisées ne doivent pas faire l’une des choses suivantes :

- Insérer, supprimer ou mettre en forme des cellules dans la feuille de calcul.
- Modifiez la valeur d’une autre cellule.
- Déplacer, renommer, supprimer ou ajouter des feuilles à un workbook.
- Modifiez l’une des options d’environnement, telles que le mode de calcul ou les affichages d’écran.
- Ajoutez des noms à un workbook.
- Définissez des propriétés ou exécutez la plupart des méthodes.

La modification d’Excel peut entraîner des performances médiocres, des dépassements de délai et des boucles infinies. Les calculs de fonction personnalisée ne doivent pas s’exécuter pendant un recalcul Excel, car ils entraînent des résultats imprévisibles.

A la place, a apporter des modifications à Excel à partir du contexte d’un bouton de ruban ou d’un volet De tâches.

## <a name="next-steps"></a>Étapes suivantes

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>Voir aussi

- [Partager des données et des événements entre les fonctions personnalisées Excel et le didacticiel du volet Des tâches](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
