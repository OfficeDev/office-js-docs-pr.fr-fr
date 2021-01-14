---
ms.date: 12/21/2020
description: Découvrez comment utiliser différents paramètres dans vos fonctions personnalisées, telles que les plages Excel, les paramètres facultatifs, le contexte d’appel, et bien plus encore.
title: Options pour les fonctions personnalisées Excel
localization_priority: Normal
ms.openlocfilehash: 312046551236e96e67de6f63f3e3511aba6f50ce
ms.sourcegitcommit: 48b9c3b63668b2a53ce73f92ce124ca07c5ca68c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/28/2020
ms.locfileid: "49735528"
---
# <a name="custom-functions-parameter-options"></a>Options des paramètres de fonctions personnalisées

Les fonctions personnalisées peuvent être configurées avec de nombreuses options de paramètres différentes.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a>Paramètres facultatifs

Lorsqu’un utilisateur appelle une fonction dans Excel, les paramètres facultatifs apparaissent entre parenthèses. Dans l’exemple suivant, la fonction Add peut éventuellement ajouter un troisième nombre. Cette fonction apparaît sous la forme `=CONTOSO.ADD(first, second, [third])` dans Excel.

#### <a name="javascript"></a>[JavaScript](#tab/javascript)

```js
/**
 * Calculates the sum of the specified numbers
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @param {number} [third] Third number to add. If omitted, third = 0.
 * @returns {number} The sum of the numbers.
 */
function add(first, second, third) {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
```

#### <a name="typescript"></a>[TypeScript](#tab/typescript)

```typescript
/**
 * Calculates the sum of the specified numbers
 * @customfunction
 * @param first First number.
 * @param second Second number.
 * @param [third] Third number to add. If omitted, third = 0.
 * @returns The sum of the numbers.
 */
function add(first: number, second: number, third?: number): number {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
```

---

> [!NOTE]
> Lorsqu’aucune valeur n’est spécifiée pour un paramètre facultatif, Excel lui affecte la valeur `null` . Cela signifie que les paramètres initialisés par défaut dans la machine à écrire ne fonctionnent pas comme prévu. N’utilisez pas la syntaxe `function add(first:number, second:number, third=0):number` car elle ne peut pas `third` être initialisée à 0. À la place, utilisez la syntaxe de la machine à écrire comme indiqué dans l’exemple précédent.

Lorsque vous définissez une fonction qui contient un ou plusieurs paramètres facultatifs, spécifiez ce qui se passe lorsque les paramètres facultatifs sont null. Dans l’exemple suivant, `zipCode` et `dayOfWeek` sont deux paramètres facultatifs pour la fonction`getWeatherReport`. Si le `zipCode` paramètre est null, la valeur par défaut est définie sur `98052` . Si le `dayOfWeek` paramètre est null, il est défini sur mercredi.

#### <a name="javascript"></a>[JavaScript](#tab/javascript)

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} [zipCode] Zip code. If omitted, zipCode = 98052.
 * @param {string} [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek) {
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

#### <a name="typescript"></a>[TypeScript](#tab/typescript)

```typescript
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param zipCode Zip code. If omitted, zipCode = 98052.
 * @param [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode?: number, dayOfWeek?: string): string {
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

---

## <a name="range-parameters"></a>Paramètres de plage

Votre fonction personnalisée peut accepter une plage de données de cellule comme paramètre d’entrée. Une fonction peut également renvoyer une plage de données. Excel passe une plage de données de cellule sous la forme d’un tableau à deux dimensions.

Par exemple, supposons que votre fonction renvoie la seconde valeur la plus élevée à partir d’une plage de nombres stockés dans Excel. La fonction suivante accepte le paramètre `values` , et la syntaxe JSDOC `number[][]` définit la propriété du paramètre `dimensionality` sur `matrix` dans les métadonnées JSON pour cette fonction. 

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.
 */
function secondHighest(values) {
  let highest = values[0][0],
    secondHighest = values[0][0];
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] >= highest) {
        secondHighest = highest;
        highest = values[i][j];
      } else if (values[i][j] >= secondHighest) {
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## <a name="repeating-parameters"></a>Paramètres répétitifs

Un paramètre extensible permet à un utilisateur d’entrer une série d’arguments facultatifs dans une fonction. Lorsque la fonction est appelée, les valeurs sont fournies dans un tableau pour le paramètre. Si le nom du paramètre se termine par un nombre, le numéro de chaque argument augmente de manière incrémentielle, par exemple `ADD(number1, [number2], [number3],…)` . Cela correspond à la Convention utilisée pour les fonctions Excel intégrées.

La fonction suivante additionne le total des nombres, des adresses de cellules, ainsi que des plages, si elles sont entrées.

```TS
/**
* The sum of all of the numbers.
* @customfunction
* @param operands A number (such as 1 or 3.1415), a cell address (such as A1 or $E$11), or a range of cell addresses (such as B3:F12)
*/

function ADD(operands: number[][][]): number {
  let total: number = 0;

  operands.forEach(range => {
    range.forEach(row => {
      row.forEach(num => {
        total += num;
      });
    });
  });

  return total;
}
```

Cette fonction s’affiche `=CONTOSO.ADD([operands], [operands]...)` dans le classeur Excel.

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a>Paramètre extensible à valeur unique

Un paramètre de valeur unique extensible permet de transmettre plusieurs valeurs uniques. Par exemple, l’utilisateur peut entrer ADD (1, B2, 3). L’exemple suivant montre comment déclarer un seul paramètre de valeur.

```JS
/**
 * @customfunction
 * @param {number[]} singleValue An array of numbers that are repeating parameters.
 */
function addSingleValue(singleValue) {
  let total = 0;
  singleValue.forEach(value => {
    total += value;
  })

  return total;
}
```

### <a name="single-range-parameter"></a>Paramètre de plage unique

Un paramètre de plage unique n’est pas techniquement un paramètre répétitif, mais il est inclus ici, car la déclaration est très similaire aux paramètres répétitifs. Il apparaîtrait à l’utilisateur sous la forme ADD (a2 : B3) où une seule plage est passée d’Excel. L’exemple suivant montre comment déclarer un paramètre de plage unique.

```JS
/**
 * @customfunction
 * @param {number[][]} singleRange
 */
function addSingleRange(singleRange) {
  let total = 0;
  singleRange.forEach(setOfSingleValues => {
    setOfSingleValues.forEach(value => {
      total += value;
    })
  })
  return total;
}
```

### <a name="repeating-range-parameter"></a>Paramètre de plage extensible

Un paramètre de plage extensible permet de transmettre plusieurs plages ou nombres. Par exemple, l’utilisateur peut entrer ADD (5, B2, C3, 8, E5 : E8). Les plages extensibles sont généralement spécifiées avec le type `number[][][]` comme il s’agit de matrices en trois dimensions. Pour un exemple, reportez-vous à l’exemple principal ci-dessous pour les paramètres de répétition (paramètres #repeating).


### <a name="declaring-repeating-parameters"></a>Déclaration de paramètres répétitifs
Dans la machine à écrire, indiquez que le paramètre est à plusieurs dimensions. Par exemple,  `ADD(values: number[])` un tableau à une dimension indiquerait `ADD(values:number[][])` un tableau à deux dimensions, et ainsi de suite.

En JavaScript, utilisez `@param values {number[]}` pour les tableaux à une dimension, `@param <name> {number[][]}` pour les tableaux à deux dimensions, et ainsi de suite pour d’autres dimensions.

Pour le format JSON dynamique, vérifiez que votre paramètre est spécifié en tant que `"repeating": true` dans votre fichier JSON, et vérifiez que vos paramètres sont marqués comme `"dimensionality": matrix` .

## <a name="invocation-parameter"></a>Paramètre invocation

Chaque fonction personnalisée reçoit automatiquement un `invocation` argument en tant que dernier paramètre d’entrée, même si elle n’est pas explicitement déclarée. Ce `invocation` paramètre correspond à l’objet [invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) . L' `Invocation` objet peut être utilisé pour récupérer un contexte supplémentaire, comme l’adresse de la cellule qui a appelé votre fonction personnalisée. Pour accéder à l' `Invocation` objet, vous devez déclarer `invocation` le dernier paramètre de votre fonction personnalisée. 

> [!NOTE]
> Le `invocation` paramètre n’apparaît pas en tant qu’argument de fonction personnalisée pour les utilisateurs dans Excel.

L’exemple suivant montre comment utiliser le `invocation` paramètre pour renvoyer l’adresse de la cellule qui a appelé votre fonction personnalisée. Cet exemple utilise la propriété [Address](/javascript/api/custom-functions-runtime/customfunctions.invocation#address) de l' `Invocation` objet. Pour accéder à l' `Invocation` objet, déclarez tout d’abord `CustomFunctions.Invocation` en tant que paramètre dans votre JSDoc. Ensuite, déclarez `@requiresAddress` dans votre JSDoc pour accéder à la `address` propriété de l' `Invocation` objet. Enfin, dans la fonction, récupérez et renvoyez la `address` propriété. 

```js
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
function getAddress(first, second, invocation) {
  var address = invocation.address;
  return address;
}
```

Dans Excel, une fonction personnalisée qui appelle la `address` propriété de l' `Invocation` objet renvoie l’adresse absolue suivant le format `SheetName!RelativeCellAddress` dans la cellule qui a appelé la fonction. Par exemple, si le paramètre d’entrée se trouve sur une feuille appelée **prix** dans la cellule F6, la valeur de l’adresse du paramètre renvoyé sera `Prices!F6` . 

Le `invocation` paramètre peut également être utilisé pour envoyer des informations à Excel. Pour en savoir plus, consultez [la rubrique créer une fonction de diffusion en continu](custom-functions-web-reqs.md#make-a-streaming-function) .

## <a name="detect-the-address-of-a-parameter"></a>Détection de l’adresse d’un paramètre

En combinaison avec le [paramètre invocation](#invocation-parameter), vous pouvez utiliser l’objet [invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) pour récupérer l’adresse d’un paramètre d’entrée de fonction personnalisée. Lorsqu’elle est appelée, la propriété [parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#parameterAddresses) de l' `Invocation` objet permet à une fonction de renvoyer les adresses de tous les paramètres d’entrée. 

Cela est utile dans les scénarios où les types de données d’entrée peuvent varier. L’adresse d’un paramètre d’entrée peut être utilisée pour vérifier le format numérique de la valeur d’entrée. Le format de nombre peut ensuite être ajusté avant l’entrée, si nécessaire. L’adresse d’un paramètre d’entrée peut également être utilisée pour détecter si la valeur d’entrée possède des propriétés connexes susceptibles de concerner les calculs ultérieurs. 

>[!IMPORTANT]
> La `parameterAddresses` propriété ne fonctionne actuellement qu’avec des [métadonnées JSON créées manuellement](custom-functions-json.md). Pour renvoyer des adresses de paramètres, la `options` propriété de l’objet doit être `requiresParameterAddresses` définie sur `true` , et l' `result` objet doit avoir la `dimensionality` propriété définie sur `matrix` .

La fonction personnalisée suivante accepte trois paramètres d’entrée, récupère la `parameterAddresses` propriété de l' `Invocation` objet pour chaque paramètre, puis renvoie ces adresses. 

```js
/**
 * Return the address of three parameters. 
 * @customfunction
 * @param {string} firstParameter First parameter.
 * @param {string} secondParameter Second parameter.
 * @param {string} thirdParameter Third parameter
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresParameterAddresses
 */
function getParameterAddresses(firstParameter, secondParameter, thirdParameter, invocation) {
  var addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

Lorsqu’une fonction personnalisée qui appelle la `parameterAddresses` propriété est exécutée, l’adresse du paramètre est renvoyée suivant le format `SheetName!RelativeCellAddress` dans la cellule qui a appelé la fonction. Par exemple, si le paramètre d’entrée se trouve sur une feuille appelée **coûts** dans la cellule D8, la valeur de l’adresse du paramètre renvoyé sera `Costs!D8` . Si la fonction personnalisée possède plusieurs paramètres et que plusieurs adresses de paramètres sont renvoyées, les adresses renvoyées s’affichent dans plusieurs cellules, décroissant verticalement, à partir de la cellule qui a appelé la fonction. 

## <a name="next-steps"></a>Étapes suivantes

Découvrez comment utiliser [des valeurs volatiles dans vos fonctions personnalisées](custom-functions-volatile.md).

## <a name="see-also"></a>Voir aussi

* [Recevoir et gérer des données à l’aide de fonctions personnalisées](custom-functions-web-reqs.md)
* [Générer automatiquement des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md)
* [Créer manuellement des métadonnées JSON pour les fonctions personnalisées](custom-functions-json.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
