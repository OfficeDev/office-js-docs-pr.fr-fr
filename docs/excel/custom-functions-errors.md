---
title: Gérer et renvoyer des erreurs à partir de votre fonction personnalisée
description: 'Gérer et retourner des erreurs comme #NULL! à partir de votre fonction personnalisée.'
ms.date: 08/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 011868f35c656869ae75c7ffab195db18f690f4f
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150344"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a>Gérer et renvoyer des erreurs à partir de votre fonction personnalisée

En cas de problème pendant l’utilisation de votre fonction personnalisée, renvoyez une erreur pour informer l’utilisateur. Si vous avez des exigences spécifiques en matière de paramètres, telles que des nombres positifs uniquement, testez les paramètres et lancez une erreur s’ils ne sont pas corrects. Vous pouvez également utiliser un bloc pour capturer les erreurs qui se produisent pendant [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) l’utilisation de votre fonction personnalisée.

## <a name="detect-and-throw-an-error"></a>Détecter et générer une erreur

Examinons un cas où vous devez vous assurer qu’un paramètre de code postal est dans le bon format pour que la fonction personnalisée fonctionne. La fonction personnalisée suivante utilise une expression régulière pour vérifier le code postal. Si le format du code postal est correct, il recherche la ville à l’aide d’une autre fonction et retourne la valeur. Si le format n’est pas valide, la fonction renvoie une `#VALUE!` erreur à la cellule.

```typescript
/**
* Gets a city name for the given U.S. zip code.
* @customfunction
* @param {string} zipCode
* @returns The city of the zip code.
*/
function getCity(zipCode: string): string {
  let isValidZip = /(^\d{5}$)|(^\d{5}-\d{4}$)/.test(zipCode);
  if (isValidZip) return cityLookup(zipCode);
  let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "Please provide a valid U.S. zip code.");
  throw error;
}
```

## <a name="the-customfunctionserror-object"></a>Objet CustomFunctions.Error

[L’objet CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) est utilisé pour renvoyer une erreur à la cellule. Lorsque vous créez l’objet, spécifiez l’erreur à utiliser en choisissant l’une des valeurs `ErrorCode` d’enum suivantes.

|Valeur enum ErrorCode  |Valeur de la cellule Excel  |Description  |
|---------------|---------|---------|
|`divisionByZero` | `#DIV/0`  | La fonction tente de diviser par zéro. |
|`invalidName`    | `#NAME?`  | Il existe une faute de frappe dans le nom de la fonction. Notez que cette erreur est prise en charge en tant qu’erreur d’entrée de fonction personnalisée, mais pas en tant qu’erreur de sortie de fonction personnalisée. |
|`invalidNumber`  | `#NUM!`   | Il y a un problème avec un nombre dans la formule. |
|`invalidReference` | `#REF!` | La fonction fait référence à une cellule non valide. Notez que cette erreur est prise en charge en tant qu’erreur d’entrée de fonction personnalisée, mais pas en tant qu’erreur de sortie de fonction personnalisée.|
|`invalidValue`   | `#VALUE!` | Une valeur dans la formule n’est pas du type. |
|`notAvailable`   | `#N/A`    | La fonction ou le service n’est pas disponible. |
|`nullReference`  | `#NULL!`  | Les plages de la formule ne se coupent pas. |

L’exemple de code suivant montre comment créer et retourner une erreur pour un nombre non valide (`#NUM!`).

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

Les `#VALUE!` `#N/A` erreurs et les erreurs sont également des messages d’erreur personnalisés. Les messages d’erreur personnalisés sont affichés dans le menu indicateur d’erreur, accessible en pointant sur l’indicateur d’erreur sur chaque cellule avec une erreur. L’exemple suivant montre comment renvoyer un message d’erreur personnalisé avec `#VALUE!` l’erreur.

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

### <a name="handle-errors-when-working-with-dynamic-arrays"></a>Gérer les erreurs lorsque vous travaillez avec des tableaux dynamiques

En plus de renvoyer une seule erreur, une fonction personnalisée peut créer un tableau dynamique qui inclut une erreur. Par exemple, une fonction personnalisée peut créer le `[1],[#NUM!],[3]` tableau. L’exemple de code suivant montre comment entrer trois paramètres dans une fonction personnalisée, remplacer l’un des paramètres d’entrée par une erreur, puis renvoyer un tableau à deux dimensions avec les résultats du traitement de chaque paramètre `#NUM!` d’entrée.

```js
/**
* Returns the #NUM! error as part of a 2-dimensional array.
* @customfunction
* @param {number} first First parameter.
* @param {number} second Second parameter.
* @param {number} third Third parameter.
* @returns {number[][]} Three results, as a 2-dimensional array.
*/
function returnInvalidNumberError(first, second, third) {
  // Use the `CustomFunctions.Error` object to retrieve an invalid number error.
  var error = new CustomFunctions.Error(
    CustomFunctions.ErrorCode.invalidNumber, // Corresponds to the #NUM! error in the Excel UI.
  );

  // Enter logic that processes the first, second, and third input parameters.
  // Imagine that the second calculation results in an invalid number error. 
  var firstResult = first;
  var secondResult =  error;
  var thirdResult = third;

  // Return the results of the first and third parameter calculations and a #NUM! error in place of the second result. 
  return [[firstResult], [secondResult], [thirdResult]];
}
```

### <a name="errors-as-custom-function-inputs"></a>Erreurs en tant qu’entrées de fonction personnalisée

Une fonction personnalisée peut être évaluée même si la plage d’entrée contient une erreur. Par exemple, une fonction personnalisée peut prendre la plage **A2:A7** comme entrée, même si **A6:A7** contient une erreur.

Pour traiter les entrées qui contiennent des erreurs, une fonction personnalisée doit avoir la propriété de métadonnées JSON `allowErrorForDataTypeAny` définie sur `true` . Pour [plus d’informations, voir Créer manuellement des métadonnées JSON pour les fonctions](custom-functions-json.md#metadata-reference) personnalisées.

> [!IMPORTANT]
> La `allowErrorForDataTypeAny` propriété ne peut être utilisée qu’avec des [métadonnées JSON créées manuellement.](custom-functions-json.md) Cette propriété ne fonctionne pas avec le processus de métadonnées JSON automatiquementgentées.

## <a name="use-trycatch-blocks"></a>Utiliser des `try...catch` blocs

En règle générale, utilisez [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) des blocs dans votre fonction personnalisée pour capturer les erreurs potentielles qui se produisent. Si vous ne traitez pas les exceptions dans votre code, elles sont renvoyées à Excel. Par défaut, Excel renvoie `#VALUE!` les erreurs ou les exceptions nonhandées.

Dans l’exemple de code suivant, la fonction personnalisée effectue un appel d’extraction à un service REST. Il est possible que l’appel échoue, par exemple, si le service REST retourne une erreur ou si le réseau est défaillant. Si cela se produit, la fonction personnalisée revient `#N/A` pour indiquer que l’appel web a échoué.

```typescript
/**
 * Gets a comment from the hypothetical contoso.com/comments API.
 * @customfunction
 * @param {number} commentID ID of a comment.
 */
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + commentID;
  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then(function (json) {
      return json.body;
    })
    .catch(function (error) {
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable);
    })
}
```

## <a name="next-steps"></a>Étapes suivantes

Découvrez comment [résoudre les problèmes liés à vos fonctions personnalisées](custom-functions-troubleshooting.md).

## <a name="see-also"></a>Voir aussi

* [Débogage des fonctions personnalisées](custom-functions-debugging.md)
* [Configuration requise de fonctions personnalisées](custom-functions-requirement-sets.md)
* [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md)
