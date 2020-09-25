---
ms.date: 09/23/2020
description: 'Gérer et retourner des erreurs comme #NULL! à partir de votre fonction personnalisée.'
title: Gérer et renvoyer des erreurs à partir de votre fonction personnalisée
localization_priority: Normal
ms.openlocfilehash: b3d3b325649a0775d3375c9f5285bba7cde0aa16
ms.sourcegitcommit: 09e1d8ff14b3c09a3eb11c91432c224a539181a4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/25/2020
ms.locfileid: "48268543"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a>Gérer et renvoyer des erreurs à partir de votre fonction personnalisée

Si un problème se présente lors de l’exécution de votre fonction personnalisée, renvoyez une erreur pour informer l’utilisateur. Si vous avez des exigences de paramètres spécifiques, telles que des nombres positifs, testez les paramètres et générez une erreur s’ils ne sont pas corrects. Vous pouvez également utiliser un bloc `try`-`catch` pour détecter les erreurs qui se produisent pendant que votre fonction personnalisée s’exécute.

## <a name="detect-and-throw-an-error"></a>Détecter et générer une erreur

Examinons un cas où vous devez vous assurer qu’un paramètre de code postal est dans le format correct pour que la fonction personnalisée fonctionne. La fonction personnalisée suivante utilise une expression régulière pour vérifier le code postal. Si le format de code postal est correct, il recherche la ville à l’aide d’une autre fonction et renvoie la valeur. Si le format n’est pas valide, la fonction renvoie une `#VALUE!` erreur à la cellule.

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

L’objet [CustomFunctions. Error](/javascript/api/custom-functions-runtime/customfunctions.error) est utilisé pour renvoyer une erreur à la cellule. Lorsque vous créez l’objet, spécifiez l’erreur que vous souhaitez utiliser en choisissant l’une des `ErrorCode` valeurs d’énumération suivantes.


|Valeur enum ErrorCode  |Valeur de la cellule Excel  |Description  |
|---------------|---------|---------|
|`divisionByZero` | `#DIV/0`  | La fonction tente d’effectuer une division par zéro. |
|`invalidName`    | `#NAME?`  | Il y a une faute de frappe dans le nom de la fonction. Notez que cette erreur est prise en charge en tant qu’erreur d’entrée d’une fonction personnalisée, mais pas en tant qu’erreur de sortie d’une fonction personnalisée. | 
|`invalidNumber`  | `#NUM!`   | Il y a un problème avec un nombre dans la formule. |
|`invalidReference` | `#REF!` | La fonction fait référence à une cellule non valide. Notez que cette erreur est prise en charge en tant qu’erreur d’entrée d’une fonction personnalisée, mais pas en tant qu’erreur de sortie d’une fonction personnalisée.|
|`invalidValue`   | `#VALUE!` | La valeur de la formule est de type incorrect. |
|`notAvailable`   | `#N/A`    | La fonction ou le service n’est pas disponible. |
|`nullReference`  | `#NULL!`  | Les plages de la formule ne se croisent pas. |

L’exemple de code suivant montre comment créer et retourner une erreur pour un nombre non valide (`#NUM!`).

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

Les `#VALUE!` `#N/A` Erreurs et prennent également en charge les messages d’erreur personnalisés. Les messages d’erreur personnalisés s’affichent dans le menu indicateur d’erreur, accessible en plaçant le curseur sur l’indicateur d’erreur sur chaque cellule avec une erreur. L’exemple suivant montre comment renvoyer un message d’erreur personnalisé avec l' `#VALUE!` erreur.

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a>Utiliser des blocs try-catch

En règle générale, utilisez des `try` - `catch` blocs dans votre fonction personnalisée pour intercepter les erreurs potentielles qui se produisent. Si vous ne gérez pas les exceptions dans votre code, celles-ci sont retournées à Excel. Par défaut, Excel renvoie `#VALUE!` des exceptions ou des erreurs non gérées.

Dans l’exemple de code suivant, la fonction personnalisée effectue un appel d’extraction à un service REST. Il est possible que l’appel échoue, par exemple, si le service REST retourne une erreur ou si le réseau est défaillant. Dans ce cas, la fonction personnalisée renvoie `#N/A` pour indiquer que l’appel Web a échoué.


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
