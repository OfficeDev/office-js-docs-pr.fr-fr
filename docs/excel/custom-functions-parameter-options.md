---
ms.date: 03/08/2021
description: Découvrez comment utiliser différents paramètres dans vos fonctions personnalisées, tels que les plages Excel, les paramètres facultatifs, le contexte d’appel, etc.
title: Options pour Excel fonctions personnalisées
ms.localizationpriority: medium
ms.openlocfilehash: 2cc0c825932afe3a70d0f9ab6483327051c199fd
ms.sourcegitcommit: 4a7b9b9b359d51688752851bf3b41b36f95eea00
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/22/2022
ms.locfileid: "63711020"
---
# <a name="custom-functions-parameter-options"></a>Options des paramètres de fonctions personnalisées

Les fonctions personnalisées sont configurables avec de nombreuses options de paramètre différentes.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a>Paramètres facultatifs

Lorsqu’un utilisateur appelle une fonction dans Excel, les paramètres facultatifs apparaissent entre parenthèses. Dans l’exemple suivant, la fonction Add peut éventuellement ajouter un troisième nombre. Cette fonction apparaît comme dans `=CONTOSO.ADD(first, second, [third])` Excel.

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
> Lorsqu’aucune valeur n’est spécifiée pour un paramètre facultatif, Excel lui affecte la valeur `null`. Cela signifie que les paramètres initialisés par défaut dans TypeScript ne fonctionneront pas comme prévu. N’utilisez pas la syntaxe `function add(first:number, second:number, third=0):number` , car elle ne s’initialisera `third` pas sur 0. Utilisez plutôt la syntaxe TypeScript comme illustré dans l’exemple précédent.

Lorsque vous définissez une fonction qui contient un ou plusieurs paramètres facultatifs, spécifiez ce qui se produit lorsque les paramètres facultatifs sont null. Dans l’exemple suivant, `zipCode` et `dayOfWeek` sont deux paramètres facultatifs pour la fonction`getWeatherReport`. Si le `zipCode` paramètre est null, la valeur par défaut est définie sur `98052`. Si le `dayOfWeek` paramètre est null, il est paramétrable mercredi.

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

Par exemple, supposons que votre fonction renvoie la seconde valeur la plus élevée à partir d’une plage de nombres stockés dans Excel. La fonction suivante accepte le `values`paramètre et la syntaxe `number[][]` `dimensionality` `matrix` JSDOC définit la propriété du paramètre dans les métadonnées JSON pour cette fonction. 

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

Un paramètre exercissable permet à un utilisateur d’entrer une série d’arguments facultatifs à une fonction. Lorsque la fonction est appelée, les valeurs sont fournies dans un tableau pour le paramètre. Si le nom du paramètre se termine par un nombre, le nombre de chaque argument augmente de manière incrémentielle, par exemple `ADD(number1, [number2], [number3],…)`. Cela correspond à la convention utilisée pour les fonctions Excel intégrées.

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

Cette fonction apparaît `=CONTOSO.ADD([operands], [operands]...)` dans le Excel de travail.

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

Un paramètre de plage unique n’est techniquement pas un paramètre exercissable, mais il est inclus ici, car la déclaration est très similaire aux paramètres exex r us. Il apparaît à l’utilisateur comme ADD(A2:B3) où une seule plage est transmise à partir Excel. L’exemple suivant montre comment déclarer un paramètre de plage unique.

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

Un paramètre de plage exercidable permet de passer plusieurs plages ou nombres. Par exemple, l’utilisateur peut entrer ADD(5,B2,C3,8,E5:E8). Les plages exercidées sont généralement spécifiées avec le type `number[][][]` , car il s’agit de matrices en trois dimensions. Pour obtenir un exemple, consultez le principal exemple répertorié pour [les paramètres répétés](#repeating-parameters).


### <a name="declaring-repeating-parameters"></a>Déclaration de paramètres répétés
Dans Typescript, indiquez que le paramètre est multidimensionnel. Par exemple,  `ADD(values: number[])` indiquerait un tableau à une dimension, `ADD(values:number[][])` un tableau à deux dimensions, etc.

Dans JavaScript, utilisez `@param values {number[]}` pour les tableaux à une dimension, `@param <name> {number[][]}` pour les tableaux à deux dimensions, et ainsi de suite pour plus de dimensions.

Pour JSON écrit à la main, `"repeating": true` assurez-vous que votre paramètre est spécifié comme dans votre fichier JSON, et vérifiez que vos paramètres sont marqués comme `"dimensionality": matrix`.

## <a name="invocation-parameter"></a>Paramètre d’appel

Chaque fonction personnalisée est automatiquement passée un `invocation` argument comme dernier paramètre d’entrée, même s’il n’est pas explicitement déclaré. Ce `invocation` paramètre correspond à [l’objet Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) . L’objet `Invocation` peut être utilisé pour récupérer un contexte supplémentaire, tel que l’adresse de la cellule qui a appelé votre fonction personnalisée. Pour accéder à l’objet `Invocation` , vous devez déclarer `invocation` comme dernier paramètre de votre fonction personnalisée. 

> [!NOTE]
> Le `invocation` paramètre n’apparaît pas en tant qu’argument de fonction personnalisée pour les utilisateurs Excel.

L’exemple suivant montre comment utiliser le paramètre `invocation` pour renvoyer l’adresse de la cellule qui a appelé votre fonction personnalisée. Cet exemple utilise la propriété [d’adresse](/javascript/api/custom-functions-runtime/customfunctions.invocation#custom-functions-runtime-customfunctions-invocation-address-member) de l’objet `Invocation` . Pour accéder à l’objet `Invocation` , déclarez d’abord `CustomFunctions.Invocation` en tant que paramètre dans votre JSDoc. Ensuite, déclarez `@requiresAddress` dans votre JSDoc pour accéder à la `address` propriété de l’objet `Invocation` . Enfin, dans la fonction, récupérez et renvoyez la `address` propriété. 

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

Dans Excel, une fonction personnalisée appelant la propriété de l’objet retourne l’adresse absolue suivant le format `SheetName!RelativeCellAddress` dans la cellule qui a `address` `Invocation` appelé la fonction. Par exemple, si le paramètre d’entrée se trouve dans une feuille appelée **Prix** dans la cellule F6, la valeur d’adresse du paramètre renvoyé est `Prices!F6`. 

Le `invocation` paramètre peut également être utilisé pour envoyer des informations à Excel. Pour en [savoir plus, voir Faire une fonction de](custom-functions-web-reqs.md#make-a-streaming-function) diffusion en continu.

## <a name="detect-the-address-of-a-parameter"></a>Détecter l’adresse d’un paramètre

En combinaison avec le paramètre [d’appel](#invocation-parameter), vous pouvez utiliser l’objet [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) pour récupérer l’adresse d’un paramètre d’entrée de fonction personnalisée. Lorsqu’elle est invoquée, [la propriété parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#custom-functions-runtime-customfunctions-invocation-parameteraddresses-member) `Invocation` de l’objet permet à une fonction de renvoyer les adresses de tous les paramètres d’entrée. 

Cela est utile dans les scénarios où les types de données d’entrée peuvent varier. L’adresse d’un paramètre d’entrée peut être utilisée pour vérifier le format numérique de la valeur d’entrée. Le format numérique peut ensuite être ajusté avant l’entrée, si nécessaire. L’adresse d’un paramètre d’entrée peut également être utilisée pour détecter si la valeur d’entrée possède des propriétés associées qui peuvent être pertinentes pour les calculs suivants. 

>[!NOTE]
> Si vous travaillez avec des métadonnées [JSON](custom-functions-json.md) créées manuellement pour renvoyer des adresses de paramètre au lieu du générateur [Yeoman pour les modules](../develop/yeoman-generator-overview.md) de Office, `options` `requiresParameterAddresses` `true`l’objet doit avoir la propriété définie sur , `result` `dimensionality` `matrix`et l’objet doit avoir la propriété définie sur .

La fonction personnalisée suivante prend trois paramètres d’entrée, `parameterAddresses` `Invocation` récupère la propriété de l’objet pour chaque paramètre, puis renvoie les adresses. 

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
  var addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

Lorsqu’une fonction personnalisée appelant la `parameterAddresses` propriété s’exécute, l’adresse du paramètre est renvoyée en suivant le format `SheetName!RelativeCellAddress` de la cellule qui a appelé la fonction. Par exemple, si le paramètre d’entrée se trouve dans une feuille appelée **Costs** dans la cellule D8, la valeur d’adresse du paramètre renvoyé est `Costs!D8`. Si la fonction personnalisée possède plusieurs paramètres et que plusieurs adresses de paramètre sont renvoyées, les adresses renvoyées se renverront sur plusieurs cellules, décroit verticalement à partir de la cellule qui a appelé la fonction. 

## <a name="next-steps"></a>Prochaines étapes

Découvrez comment utiliser des [valeurs volatiles dans vos fonctions personnalisées](custom-functions-volatile.md).

## <a name="see-also"></a>Voir aussi

* [Recevoir et gérer des données à l’aide de fonctions personnalisées](custom-functions-web-reqs.md)
* [Générer automatiquement des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md)
* [Créer manuellement des métadonnées JSON pour les fonctions personnalisées](custom-functions-json.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
