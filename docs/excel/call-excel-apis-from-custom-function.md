---
title: Appeler Excel API JavaScript à partir d’une fonction personnalisée
description: Découvrez les Excel JavaScript que vous pouvez appeler à partir de votre fonction personnalisée.
ms.date: 03/05/2021
localization_priority: Normal
ms.openlocfilehash: d44f88dc11136bd0302453054cefe93c82b22136e2084baecac006834100a077
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57079848"
---
# <a name="call-excel-javascript-apis-from-a-custom-function"></a>Appeler Excel API JavaScript à partir d’une fonction personnalisée

Appelez Excel API JavaScript à partir de vos fonctions personnalisées pour obtenir des données de plage et obtenir plus de contexte pour vos calculs. Appeler Excel API JavaScript par le biais d’une fonction personnalisée peut être utile lorsque :

- Une fonction personnalisée doit obtenir des informations de la Excel avant le calcul. Ces informations peuvent inclure des propriétés de document, des formats de plage, des parties XML personnalisées, un nom de Excel informations spécifiques.
- Une fonction personnalisée définira le format numérique de la cellule pour les valeurs de retour après le calcul.

> [!IMPORTANT]
> Pour appeler Excel API JavaScript à partir de votre fonction personnalisée, vous devez utiliser un runtime JavaScript partagé. Pour plus d’information, consultez [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="code-sample"></a>Exemple de code

Pour appeler Excel API JavaScript à partir d’une fonction personnalisée, vous avez d’abord besoin d’un contexte. Utilisez le [Excel. Objet RequestContext](/javascript/api/excel/excel.requestcontext) pour obtenir un contexte. Utilisez ensuite le contexte pour appeler les API dont vous avez besoin dans le workbook.

L’exemple de code suivant montre comment utiliser pour obtenir une valeur à partir `Excel.RequestContext` d’une cellule dans le workbook. Dans cet exemple, le paramètre est transmis à la Excel `address` de l’API JavaScript [Worksheet.getRange](/javascript/api/excel/excel.worksheet#getRange_address_) et doit être entré sous forme de chaîne. Par exemple, la fonction personnalisée entrée dans l’interface utilisateur Excel doit suivre le modèle , où est l’adresse de la cellule à partir de laquelle récupérer `=CONTOSO.GETRANGEVALUE("A1")` `"A1"` la valeur.

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

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a>Limitations de l’appel Excel api JavaScript par le biais d’une fonction personnalisée

N’appelez pas Excel API JavaScript à partir d’une fonction personnalisée qui modifie l’environnement de Excel. Cela signifie que vos fonctions personnalisées ne doivent pas faire l’une des choses suivantes :

- Insérer, supprimer ou mettre en forme des cellules dans la feuille de calcul.
- Modifiez la valeur d’une autre cellule.
- Déplacer, renommer, supprimer ou ajouter des feuilles à un workbook.
- Modifiez l’une des options d’environnement, telles que le mode de calcul ou les affichages d’écran.
- Ajoutez des noms à un workbook.
- Définissez des propriétés ou exécutez la plupart des méthodes.

Le Excel peut entraîner des performances médiocres, des délai d’exécution et des boucles infinies. Les calculs de fonction personnalisée ne doivent pas s’exécuter pendant qu’Excel recalcul est en cours, car il se traduit par des résultats imprévisibles.

A la place, a apporté des modifications Excel à partir du contexte d’un bouton de ruban ou d’un volet De tâches.

## <a name="next-steps"></a>Étapes suivantes

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>Voir aussi

- [Partager des données et des événements entre Excel fonctions personnalisées et didacticiel du volet Des tâches](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Configurer votre complément Office pour utiliser un runtime JavaScript partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
