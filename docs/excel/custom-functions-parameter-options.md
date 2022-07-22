---
title: Options pour les fonctions personnalisées Excel
description: Découvrez comment utiliser différents paramètres dans vos fonctions personnalisées, telles que les plages Excel, les paramètres facultatifs, le contexte d’appel, etc.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: de86afc60d7d0b81820bd742e989e0ee7dd6970c
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958572"
---
# <a name="custom-functions-parameter-options"></a>Options des paramètres des fonctions personnalisées

Les fonctions personnalisées sont configurables avec de nombreuses options de paramètres différentes.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a>Paramètres facultatifs

Lorsqu’un utilisateur appelle une fonction dans Excel, les paramètres facultatifs apparaissent entre parenthèses. Dans l’exemple suivant, la fonction add peut éventuellement ajouter un troisième nombre. Cette fonction apparaît comme `=CONTOSO.ADD(first, second, [third])` dans Excel.

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
> Lorsqu’aucune valeur n’est spécifiée pour un paramètre facultatif, Excel lui attribue la valeur `null`. Cela signifie que les paramètres initialisés par défaut dans TypeScript ne fonctionneront pas comme prévu. N’utilisez pas la syntaxe `function add(first:number, second:number, third=0):number` , car elle n’initialisera `third` pas sur 0. Utilisez plutôt la syntaxe TypeScript comme indiqué dans l’exemple précédent.

Lorsque vous définissez une fonction qui contient un ou plusieurs paramètres facultatifs, spécifiez ce qui se passe lorsque les paramètres facultatifs sont null. Dans l’exemple suivant, `zipCode` et `dayOfWeek` sont deux paramètres facultatifs pour la fonction`getWeatherReport`. Si le `zipCode` paramètre est null, la valeur par défaut est définie sur `98052`. Si le `dayOfWeek` paramètre est null, il est défini sur mercredi.

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

Votre fonction personnalisée peut accepter une plage de données de cellule comme paramètre d’entrée. Une fonction peut également retourner une plage de données. Excel transmet une plage de données de cellule sous forme de tableau à deux dimensions.

Par exemple, supposons que votre fonction renvoie la seconde valeur la plus élevée à partir d’une plage de nombres stockés dans Excel. La fonction suivante accepte le paramètre `values`et la syntaxe `number[][]` JSDOC définit la propriété `matrix` du `dimensionality` paramètre sur les métadonnées JSON pour cette fonction.

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.
 */
function secondHighest(values) {
  let highest = values[0][0],
    secondHighest = values[0][0];
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
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

Un paramètre répétitif permet à un utilisateur d’entrer une série d’arguments facultatifs dans une fonction. Lorsque la fonction est appelée, les valeurs sont fournies dans un tableau pour le paramètre. Si le nom du paramètre se termine par un nombre, le nombre de chaque argument augmente de façon incrémentielle, par `ADD(number1, [number2], [number3],…)`exemple . Cela correspond à la convention utilisée pour les fonctions Excel intégrées.

La fonction suivante additionnera le total des nombres, des adresses de cellule et des plages, s’il est entré.

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

![Fonction personnalisée ADD entrée dans la cellule d’une feuille de calcul Excel](../images/operands.png)

### <a name="repeating-single-value-parameter"></a>Répétition d’un paramètre de valeur unique

Un paramètre de valeur unique répété permet de passer plusieurs valeurs uniques. Par exemple, l’utilisateur peut entrer ADD(1,B2,3). L’exemple suivant montre comment déclarer un paramètre de valeur unique.

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

Un paramètre de plage unique n’est techniquement pas un paramètre répétitif, mais il est inclus ici, car la déclaration est très similaire aux paramètres répétitifs. Il semblerait que l’utilisateur soit ADD(A2:B3) où une plage unique est passée à partir d’Excel. L’exemple suivant montre comment déclarer un paramètre de plage unique.

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

### <a name="repeating-range-parameter"></a>Paramètre de plage répétitif

Un paramètre de plage répétée permet de passer plusieurs plages ou nombres. Par exemple, l’utilisateur peut entrer ADD(5,B2,C3,8,E5:E8). Les plages répétées sont généralement spécifiées avec le type `number[][][]` , car il s’agit de matrices à trois dimensions. Pour obtenir un exemple, consultez l’exemple principal répertorié pour [les paramètres répétés](#repeating-parameters).

### <a name="declaring-repeating-parameters"></a>Déclaration de paramètres répétitifs

Dans Typescript, indiquez que le paramètre est multidimensionnel. Par exemple,  `ADD(values: number[])` indiquerait un tableau unidimensionnel, `ADD(values:number[][])` indiquerait un tableau à deux dimensions, et ainsi de suite.

En JavaScript, utilisez `@param values {number[]}` des tableaux unidimensionnels, `@param <name> {number[][]}` des tableaux à deux dimensions, etc. pour d’autres dimensions.

Pour le code JSON créé à la main, vérifiez que votre paramètre est spécifié comme `"repeating": true` dans votre fichier JSON, et vérifiez que vos paramètres sont marqués comme `"dimensionality": matrix`.

## <a name="invocation-parameter"></a>Paramètre d’appel

Chaque fonction personnalisée reçoit automatiquement un `invocation` argument comme dernier paramètre d’entrée, même s’il n’est pas explicitement déclaré. Ce `invocation` paramètre correspond à l’objet [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) . L’objet `Invocation` peut être utilisé pour récupérer un contexte supplémentaire, tel que l’adresse de la cellule qui a appelé votre fonction personnalisée. Pour accéder à l’objet `Invocation` , vous devez déclarer `invocation` comme dernier paramètre dans votre fonction personnalisée.

> [!NOTE]
> Le `invocation` paramètre n’apparaît pas comme un argument de fonction personnalisée pour les utilisateurs dans Excel.

L’exemple suivant montre comment utiliser le `invocation` paramètre pour retourner l’adresse de la cellule qui a appelé votre fonction personnalisée. Cet exemple utilise la propriété [d’adresse](/javascript/api/custom-functions-runtime/customfunctions.invocation#custom-functions-runtime-customfunctions-invocation-address-member) de l’objet `Invocation` . Pour accéder à l’objet `Invocation` , déclarez `CustomFunctions.Invocation` d’abord en tant que paramètre dans votre JSDoc. Ensuite, déclarez `@requiresAddress` dans votre JSDoc pour accéder à la `address` propriété de l’objet `Invocation` . Enfin, dans la fonction, récupérez puis retournez la `address` propriété.

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
  const address = invocation.address;
  return address;
}
```

Dans Excel, une fonction personnalisée appelant la `address` propriété de l’objet `Invocation` retourne l’adresse absolue suivant le format `SheetName!RelativeCellAddress` de la cellule qui a appelé la fonction. Par exemple, si le paramètre d’entrée se trouve sur une feuille appelée **Prices** dans la cellule F6, la valeur d’adresse du paramètre retourné est `Prices!F6`.

Le `invocation` paramètre peut également être utilisé pour envoyer des informations à Excel. Pour en savoir plus [, consultez Créer une fonction de diffusion en continu](custom-functions-web-reqs.md#make-a-streaming-function) .

## <a name="detect-the-address-of-a-parameter"></a>Détecter l’adresse d’un paramètre

En combinaison avec le [paramètre d’appel](#invocation-parameter), vous pouvez utiliser l’objet [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) pour récupérer l’adresse d’un paramètre d’entrée de fonction personnalisée. Lorsqu’elle est appelée, la propriété [parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#custom-functions-runtime-customfunctions-invocation-parameteraddresses-member) de l’objet `Invocation` permet à une fonction de retourner les adresses de tous les paramètres d’entrée.

Cela est utile dans les scénarios où les types de données d’entrée peuvent varier. L’adresse d’un paramètre d’entrée peut être utilisée pour vérifier le format numérique de la valeur d’entrée. Le format de nombre peut ensuite être ajusté avant l’entrée, si nécessaire. L’adresse d’un paramètre d’entrée peut également être utilisée pour détecter si la valeur d’entrée a des propriétés connexes susceptibles d’être pertinentes pour les calculs suivants.

>[!NOTE]
> Si vous utilisez des [métadonnées JSON créées manuellement](custom-functions-json.md) pour retourner des adresses de paramètres au lieu du [générateur Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md), la `options` propriété doit être définie `true`sur l’objet `requiresParameterAddresses` et `result` la `dimensionality` propriété doit être définie `matrix`sur .

La fonction personnalisée suivante accepte trois paramètres d’entrée, récupère la `parameterAddresses` propriété de l’objet `Invocation` pour chaque paramètre, puis retourne les adresses.

```js
/**
 * Return the addresses of three parameters. 
 * @customfunction
 * @param {string} firstParameter First parameter.
 * @param {string} secondParameter Second parameter.
 * @param {string} thirdParameter Third parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @returns {string[][]} The addresses of the parameters, as a 2-dimensional array. 
 * @requiresParameterAddresses
 */
function getParameterAddresses(firstParameter, secondParameter, thirdParameter, invocation) {
  const addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

Lorsqu’une fonction personnalisée appelant la `parameterAddresses` propriété s’exécute, l’adresse du paramètre est retournée en suivant le format `SheetName!RelativeCellAddress` de la cellule qui a appelé la fonction. Par exemple, si le paramètre d’entrée se trouve sur une feuille appelée **Costs** dans la cellule D8, la valeur d’adresse du paramètre retourné est `Costs!D8`. Si la fonction personnalisée a plusieurs paramètres et que plusieurs adresses de paramètre sont retournées, les adresses retournées se répandent sur plusieurs cellules, descendant verticalement à partir de la cellule qui a appelé la fonction.

## <a name="next-steps"></a>Étapes suivantes

Découvrez comment utiliser [des valeurs volatiles dans vos fonctions personnalisées](custom-functions-volatile.md).

## <a name="see-also"></a>Voir aussi

- [Recevoir et gérer des données à l’aide de fonctions personnalisées](custom-functions-web-reqs.md)
- [Générer automatiquement des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md)
- [Créer manuellement des métadonnées JSON pour des fonctions personnalisées](custom-functions-json.md)
- [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
- [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
