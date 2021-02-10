---
ms.date: 02/04/2021
description: Découvrez comment utiliser différents paramètres dans vos fonctions personnalisées, tels que les plages Excel, les paramètres facultatifs, le contexte d’appel, etc.
title: Options pour les fonctions personnalisées Excel
localization_priority: Normal
ms.openlocfilehash: afe6947b1a1b9022a0284535b9ab1d68c9777c14
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173905"
---
# <a name="custom-functions-parameter-options"></a>Options des paramètres de fonctions personnalisées

Les fonctions personnalisées sont configurables avec de nombreuses options de paramètre différentes.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a>Paramètres facultatifs

Lorsqu’un utilisateur appelle une fonction dans Excel, les paramètres facultatifs apparaissent entre parenthèses. Dans l’exemple suivant, la fonction Add peut éventuellement ajouter un troisième nombre. Cette fonction apparaît comme `=CONTOSO.ADD(first, second, [third])` dans Excel.

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
> Lorsqu’aucune valeur n’est spécifiée pour un paramètre facultatif, Excel lui affecte la valeur `null` . Cela signifie que les paramètres initialisés par défaut dans TypeScript ne fonctionneront pas comme prévu. N’utilisez pas la `function add(first:number, second:number, third=0):number` syntaxe, car elle ne s’initialisera pas sur `third` 0. Utilisez plutôt la syntaxe TypeScript comme illustré dans l’exemple précédent.

Lorsque vous définissez une fonction qui contient un ou plusieurs paramètres facultatifs, spécifiez ce qui se produit lorsque les paramètres facultatifs sont null. Dans l’exemple suivant, `zipCode` et `dayOfWeek` sont deux paramètres facultatifs pour la fonction`getWeatherReport`. Si le `zipCode` paramètre est null, la valeur par défaut est définie sur `98052` . Si le `dayOfWeek` paramètre est null, il est paramétrable mercredi.

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

Votre fonction personnalisée peut accepter une plage de données de cellule comme paramètre d’entrée. Une fonction peut également renvoyer une plage de données. Excel passe une plage de données de cellule sous forme de tableau à deux dimensions.

Par exemple, supposons que votre fonction renvoie la seconde valeur la plus élevée à partir d’une plage de nombres stockés dans Excel. La fonction suivante accepte le paramètre et la syntaxe JSDOC définit la propriété du paramètre dans les métadonnées `values` `number[][]` `dimensionality` `matrix` JSON pour cette fonction. 

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

## <a name="repeating-parameters"></a>Paramètres répétés

Un paramètre exercissable permet à un utilisateur d’entrer une série d’arguments facultatifs à une fonction. Lorsque la fonction est appelée, les valeurs sont fournies dans un tableau pour le paramètre. Si le nom du paramètre se termine par un nombre, le nombre de chaque argument augmente de manière incrémentielle, par `ADD(number1, [number2], [number3],…)` exemple. Cela correspond à la convention utilisée pour les fonctions Excel intégrées.

La fonction suivante additione le total des nombres, des adresses de cellule, ainsi que des plages, si elles sont entrées.

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

Cette fonction `=CONTOSO.ADD([operands], [operands]...)` s’affiche dans le livre de calcul Excel.

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a>Paramètre de valeur unique répété

Un paramètre à valeur unique exercissable permet de passer plusieurs valeurs simples. Par exemple, l’utilisateur peut entrer ADD(1,B2,3). L’exemple suivant montre comment déclarer un paramètre de valeur unique.

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

Un paramètre de plage unique n’est techniquement pas un paramètre exercissable, mais il est inclus ici, car la déclaration est très similaire aux paramètres ext ments ex r us. Il apparaît à l’utilisateur comme ADD(A2:B3) où une seule plage est passée à partir d’Excel. L’exemple suivant montre comment déclarer un paramètre de plage unique.

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

### <a name="repeating-range-parameter"></a>Paramètre de plage répétée

Un paramètre de plage exercidable permet de passer plusieurs plages ou nombres. Par exemple, l’utilisateur peut entrer ADD(5,B2,C3,8,E5:E8). Les plages exercidées sont généralement spécifiées avec le type, car il s’agit de `number[][][]` matrices en trois dimensions. Pour obtenir un exemple, consultez le principal exemple répertorié pour les paramètres répétés(#repeating-parameters).


### <a name="declaring-repeating-parameters"></a>Déclaration de paramètres répétés
Dans Typescript, indiquez que le paramètre est multidimensionnel. Par exemple, cela indiquerait un tableau à une dimension, un tableau à  `ADD(values: number[])` `ADD(values:number[][])` deux dimensions, etc.

Dans JavaScript, utilisez pour les tableaux à une dimension, pour les tableaux à deux dimensions, et ainsi de `@param values {number[]}` suite pour plus de `@param <name> {number[][]}` dimensions.

Pour JSON écrit à la main, assurez-vous que votre paramètre est spécifié comme dans votre fichier JSON, et vérifiez que vos paramètres sont `"repeating": true` marqués comme `"dimensionality": matrix` .

## <a name="invocation-parameter"></a>Paramètre d’appel

Chaque fonction personnalisée est automatiquement passée un argument comme dernier paramètre `invocation` d’entrée, même s’il n’est pas explicitement déclaré. Ce `invocation` paramètre correspond à l’objet [Invocation.](/javascript/api/custom-functions-runtime/customfunctions.invocation) L’objet peut être utilisé pour récupérer un contexte supplémentaire, tel que l’adresse de la cellule `Invocation` qui a appelé votre fonction personnalisée. Pour accéder à `Invocation` l’objet, vous devez déclarer `invocation` comme dernier paramètre de votre fonction personnalisée. 

> [!NOTE]
> Le `invocation` paramètre n’apparaît pas en tant qu’argument de fonction personnalisée pour les utilisateurs dans Excel.

L’exemple suivant montre comment utiliser le paramètre pour renvoyer l’adresse de la cellule `invocation` qui a appelé votre fonction personnalisée. Cet exemple utilise la propriété [d’adresse](/javascript/api/custom-functions-runtime/customfunctions.invocation#address) de `Invocation` l’objet. Pour accéder à `Invocation` l’objet, déclarez d’abord `CustomFunctions.Invocation` en tant que paramètre dans votre JSDoc. Ensuite, déclarez `@requiresAddress` dans votre JSDoc pour accéder à la `address` propriété de `Invocation` l’objet. Enfin, dans la fonction, récupérez et renvoyez la `address` propriété. 

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

Dans Excel, une fonction personnalisée appelant la propriété de l’objet retourne l’adresse absolue en suivant le format de la cellule qui a `address` `Invocation` appelé la `SheetName!RelativeCellAddress` fonction. Par exemple, si le paramètre d’entrée se trouve dans une feuille appelée **Prix** dans la cellule F6, la valeur d’adresse du paramètre renvoyé est `Prices!F6` . 

Le `invocation` paramètre peut également être utilisé pour envoyer des informations à Excel. Pour en [savoir plus, voir](custom-functions-web-reqs.md#make-a-streaming-function) Faire une fonction de diffusion en continu.

## <a name="detect-the-address-of-a-parameter"></a>Détecter l’adresse d’un paramètre

En combinaison avec le paramètre [d’appel,](#invocation-parameter)vous pouvez utiliser l’objet [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) pour récupérer l’adresse d’un paramètre d’entrée de fonction personnalisée. Lorsqu’elle est invoquée, [la propriété parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#parameterAddresses) de l’objet permet à une fonction de renvoyer les adresses de tous `Invocation` les paramètres d’entrée. 

Cela est utile dans les scénarios où les types de données d’entrée peuvent varier. L’adresse d’un paramètre d’entrée peut être utilisée pour vérifier le format numérique de la valeur d’entrée. Le format numérique peut ensuite être ajusté avant l’entrée, si nécessaire. L’adresse d’un paramètre d’entrée peut également être utilisée pour détecter si la valeur d’entrée possède des propriétés associées qui peuvent être pertinentes pour les calculs suivants. 

>[!NOTE]
> Si vous travaillez avec des métadonnées [JSON](custom-functions-json.md) créées manuellement pour renvoyer des adresses de paramètre au lieu du générateur Yo Office, l’objet doit avoir la propriété définie sur , et l’objet doit avoir la propriété définie sur `options` `requiresParameterAddresses` `true` `result` `dimensionality` `matrix` .

La fonction personnalisée suivante prend trois paramètres d’entrée, récupère la propriété de l’objet pour chaque paramètre, puis `parameterAddresses` `Invocation` renvoie les adresses. 

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

Lorsqu’une fonction personnalisée appelant la propriété s’exécute, l’adresse du paramètre est renvoyée en suivant le format de la cellule `parameterAddresses` qui a appelé la `SheetName!RelativeCellAddress` fonction. Par exemple, si le paramètre d’entrée se trouve dans une feuille appelée **Costs** dans la cellule D8, la valeur d’adresse du paramètre renvoyé est `Costs!D8` . Si la fonction personnalisée possède plusieurs paramètres et que plusieurs adresses de paramètre sont renvoyées, les adresses renvoyées se renverront sur plusieurs cellules, décroit verticalement à partir de la cellule qui a appelé la fonction. 

## <a name="next-steps"></a>Étapes suivantes

Découvrez comment utiliser des [valeurs volatiles dans vos fonctions personnalisées.](custom-functions-volatile.md)

## <a name="see-also"></a>Voir aussi

* [Recevoir et gérer des données à l’aide de fonctions personnalisées](custom-functions-web-reqs.md)
* [Générer automatiquement des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md)
* [Créer manuellement des métadonnées JSON pour les fonctions personnalisées](custom-functions-json.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
