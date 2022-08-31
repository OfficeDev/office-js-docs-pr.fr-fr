---
title: Appeler des API JavaScript Excel à partir d’une fonction personnalisée
description: Découvrez les API JavaScript Excel que vous pouvez appeler à partir de votre fonction personnalisée.
ms.date: 08/30/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8d1cbf6d07e4ede5b8309e899828f8f1d8ad1fa0
ms.sourcegitcommit: eef2064d7966db91f8401372dd255a32d76168c2
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/31/2022
ms.locfileid: "67464831"
---
# <a name="call-excel-javascript-apis-from-a-custom-function"></a>Appeler des API JavaScript Excel à partir d’une fonction personnalisée

Appelez les API JavaScript Excel à partir de vos fonctions personnalisées pour obtenir des données de plage et obtenir plus de contexte pour vos calculs. L’appel d’API JavaScript Excel via une fonction personnalisée peut être utile dans les cas suivants :

- Une fonction personnalisée doit obtenir des informations à partir d’Excel avant le calcul. Ces informations peuvent inclure des propriétés de document, des formats de plage, des parties XML personnalisées, un nom de classeur ou d’autres informations spécifiques à Excel.
- Une fonction personnalisée définit le format numérique de la cellule pour les valeurs de retour après calcul.

> [!IMPORTANT]
> Pour appeler des API JavaScript Excel à partir de votre fonction personnalisée, vous devez utiliser un [runtime partagé](../testing/runtimes.md#shared-runtime). Pour plus d’informations, consultez [Configurer votre complément Office pour utiliser un runtime partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md) .

## <a name="code-sample"></a>Exemple de code

Pour appeler des API JavaScript Excel à partir d’une fonction personnalisée, vous avez d’abord besoin d’un contexte. Utilisez l’objet [Excel.RequestContext](/javascript/api/excel/excel.requestcontext) pour obtenir un contexte. Utilisez ensuite le contexte pour appeler les API dont vous avez besoin dans le classeur.

L’exemple de code suivant montre comment utiliser `Excel.RequestContext` pour obtenir une valeur à partir d’une cellule du classeur. Dans cet exemple, le `address` paramètre est transmis à la méthode [Worksheet.getRange](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1)) de l’API JavaScript Excel et doit être entré sous forme de chaîne. Par exemple, la fonction personnalisée entrée dans l’interface utilisateur Excel doit suivre le modèle `=CONTOSO.GETRANGEVALUE("A1")`, où `"A1"` est l’adresse de la cellule à partir de laquelle récupérer la valeur.

```JavaScript
/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
async function getRangeValue(address) {
 // Retrieve the context object. 
 const context = new Excel.RequestContext();
 
 // Use the context object to access the cell at the input address. 
 const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load("values");
 await context.sync();
 
 // Return the value of the cell at the input address.
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a>Limitations de l’appel d’API JavaScript Excel par le biais d’une fonction personnalisée

Un complément de fonctions personnalisées peut appeler des API JavaScript Excel, mais vous devez être prudent quant aux API qu’il appelle. N’appelez pas les API JavaScript Excel à partir d’une fonction personnalisée qui modifie des cellules en dehors de la cellule exécutant la fonction personnalisée. La modification d’autres cellules ou de l’environnement Excel peut entraîner des performances médiocres, des délais d’expiration et des boucles infinies dans l’application Excel. Cela signifie que vos fonctions personnalisées ne doivent pas effectuer l’une des opérations suivantes :

- Insérer, supprimer ou mettre en forme des cellules dans la feuille de calcul.
- Modifiez la valeur d’une autre cellule.
- Déplacez, renommez, supprimez ou ajoutez des feuilles à un classeur.
- Ajoutez des noms à un classeur.
- Définissez les propriétés.
- Modifiez l’une des options d’environnement Excel, telles que le mode de calcul ou les vues d’écran.

Votre complément de fonctions personnalisées peut lire des informations à partir de cellules extérieures à la cellule exécutant la fonction personnalisée, mais il ne doit pas effectuer d’opérations d’écriture dans d’autres cellules. Au lieu de cela, apportez des modifications à d’autres cellules ou à l’environnement Excel à partir du contexte d’un bouton de ruban ou d’un volet Office. En outre, les calculs de fonction personnalisés ne doivent pas s’exécuter pendant qu’un recalcul Excel est en cours, car ce scénario crée des résultats imprévisibles.

## <a name="next-steps"></a>Prochaines étapes

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>Voir aussi

- [Partager des données et des événements entre les fonctions personnalisées Excel et le didacticiel du volet Office](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Configurer votre complément Office pour utiliser un runtime partagé](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
