---
ms.date: 05/09/2019
description: Découvrez comment utiliser différents paramètres dans vos fonctions personnalisées, telles que les plages Excel, les paramètres facultatifs, le contexte d’appel, et bien plus encore.
title: Options pour les fonctions personnalisées Excel
localization_priority: Normal
ms.openlocfilehash: ba437f3a49ec3129b72f3396e85fcbd46af82cb7
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952074"
---
# <a name="custom-functions-parameter-options"></a>Options des paramètres de fonctions personnalisées

Les fonctions personnalisées peuvent être configurées avec de nombreuses options différentes pour les paramètres:
- [Paramètres facultatifs](#custom-functions-optional-parameters)
- [Paramètres de plage](#range-parameters)
- [Paramètre de contexte d’invocation](#invocation-parameter)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="custom-functions-optional-parameters"></a>Paramètres facultatifs de fonctions personnalisées

Alors que les paramètres réguliers sont obligatoires, les paramètres facultatifs ne le sont pas. Lorsqu’un utilisateur appelle une fonction dans Excel, les paramètres facultatifs apparaissent entre parenthèses. Dans l’exemple suivant, la fonction Add peut éventuellement ajouter un troisième nombre. Cette fonction apparaît sous `=CONTOSO.ADD(first, second, [third])` la forme dans Excel.

```js
/**
 * Add two numbers
 * @customfunction 
 * @param {number} first First number.
 * @param {number} second Second number.
 * @param {number} [third] Third number to add. If omitted, third = 0.
 * @returns {number} The sum of the numbers.
 */
function add(first, second, third) {
  if (third === undefined) {
    return first + second + third;
  }
  return first + second;
}
CustomFunctions.associate("ADD", add);
```

Lorsque vous définissez une fonction qui contient un ou plusieurs paramètres facultatifs, vous devez spécifier ce qu’il se passe lorsque les paramètres facultatifs ne sont pas définis. Dans l’exemple suivant, `zipCode` et `dayOfWeek` sont deux paramètres facultatifs pour la fonction`getWeatherReport`. Si le `zipCode` paramètre n’est pas défini, la valeur par défaut est définie `98052`sur. Si le paramètre`dayOfWeek` n’est pas défini, la valeur par défaut est définie à mercredi.

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} zipCode Zip code. If omitted, zipCode = 98052.
 * @param {string} dayOfWeek Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === undefined) {
      zipCode = "98052";
  }

  if (dayOfWeek === undefined) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

## <a name="range-parameters"></a>Paramètres de plage

Votre fonction personnalisée peut accepter une plage de données de cellule comme paramètre d’entrée. Une fonction peut également renvoyer une plage de données. Excel passe une plage de données de cellule sous la forme d’un tableau à deux dimensions.

Par exemple, supposons que votre fonction renvoie la seconde valeur la plus élevée à partir d’une plage de nombres stockés dans Excel. La fonction suivante prend le paramètre `values`, c’est-à-dire un type de `Excel.CustomFunctionDimensionality.matrix`. Notez que dans les métadonnées JSON pour cette fonction, la propriété `type` du paramètre est définie `matrix`sur.

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.  
 */
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
  for(var i = 0; i < values.length; i++){
    for(var j = 0; j < values[i].length; j++){
      if(values[i][j] >= highest){
        secondHighest = highest;
        highest = values[i][j];
      }
      else if(values[i][j] >= secondHighest){
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
CustomFunctions.associate("SECONDHIGHEST", secondHighest);
```

## <a name="invocation-parameter"></a>Paramètre invocation

Chaque fonction personnalisée reçoit automatiquement un `invocation` argument en tant que dernier argument. Cet argument peut être utilisé pour récupérer un contexte supplémentaire, comme l’adresse de la cellule d’appel. Ou elle peut être utilisée pour envoyer des informations à Excel, comme un gestionnaire de fonctions pour [annuler une fonction](custom-functions-web-reqs.md#stream-and-cancel-functions). Même si aucun paramètre n’est déclaré, votre fonction personnalisée a ce paramètre. Cet argument n’apparaît pas pour un utilisateur dans Excel. Si vous souhaitez utiliser `invocation` dans votre fonction personnalisée, déclarez-le comme dernier paramètre.

Dans l’exemple de code suivant, `invocation` le contexte est explicitement indiqué pour votre référence.

```js
/**
 * Add two numbers.
 * @customfunction 
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two (or optionally three) numbers.
 */
function add(first, second, invocation) {
  return first + second;
}
CustomFunctions.associate("ADD", add);
```

Le paramètre vous permet d’obtenir le contexte de la cellule d’appel, ce qui peut être utile dans certains scénarios, notamment [la découverte de l’adresse d’une cellule qui appelle une fonction personnalisée](#addressing-cells-context-parameter).

### <a name="addressing-cells-context-parameter"></a>Paramètre de contexte de la cellule d’adressage

Dans certains cas, vous devez obtenir l’adresse de la cellule qui a appelé votre fonction personnalisée. Cela est utile dans les scénarios suivants:

- Mise en forme des plages: utilisez l’adresse de la cellule comme clé pour stocker des informations dans [OfficeRuntime. Storage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data). Utilisez ensuite [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) dans Excel pour charger la clé à partir de l’élément `OfficeRuntime.storage`.
- Affichage de valeurs mises en cache : si votre fonction est utilisée en mode hors connexion, affichez les valeurs mises en cache à partir de l’élément `OfficeRuntime.storage` à l’aide de `onCalculated`.
- Rapprochement : utilisez l’adresse de la cellule pour découvrir la cellule d’origine afin de vous aider à réaliser un rapprochement lors du traitement.

Pour demander le contexte d’une cellule d’adressage dans une fonction, vous devez utiliser une fonction pour Rechercher l’adresse de la cellule, comme dans l’exemple suivant. Les informations relatives à l’adresse d’une cellule ne sont `@requiresAddress` exposées que si elles sont balisées dans les commentaires de la fonction.

```js
/**
 * Function that gets the address of a cell.
 * @customfunction
 * @param {CustomFunctions.Invocation} invocation Uses the invocation parameter present in each cell.
 * @requiresAddress
 * @returns {string} Returns address of cell.
 */

function getAddress(invocation) {
  return invocation.address;
}
CustomFunctions.associate("GETADDRESS", getAddress);
```

Par défaut, les valeurs renvoyées par une fonction `getAddress` ont le format suivant : `SheetName!CellNumber`. Par exemple, si une fonction a été appelée à partir d’une feuille de calcul appelée Dépenses dans la cellule B2, la valeur renvoyée serait `Expenses!B2`.

## <a name="next-steps"></a>Étapes suivantes
Découvrez comment [enregistrer l’État dans vos fonctions personnalisées](custom-functions-save-state.md) ou utiliser des [valeurs volatiles dans vos fonctions personnalisées](custom-functions-volatile.md).

## <a name="see-also"></a>Voir aussi

* [Recevoir et gérer des données avec des fonctions personnalisées](custom-functions-web-reqs.md)
* [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md)
* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Générer automatiquement des métadonnées JSON pour les fonctions personnalisées](custom-functions-json-autogeneration.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
