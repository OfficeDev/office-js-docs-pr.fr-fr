---
ms.date: 03/11/2020
description: 'Gérer et retourner des erreurs comme #NULL! à partir de votre fonction personnalisée'
title: Gérer et retourner des erreurs à partir de votre fonction personnalisée (préversion)
localization_priority: Normal
ms.openlocfilehash: 10bb7ca6ff612ef38b26b88fed5ce9ce81ed7edb
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717046"
---
# <a name="handle-and-return-errors-from-your-custom-function-preview"></a>Gérer et retourner des erreurs à partir de votre fonction personnalisée (préversion)

> [!NOTE]
> Les fonctionnalités décrites dans cet article sont actuellement en préversion et peuvent faire l’objet de modifications. Elles ne sont pas prises en charge dans les environnements de production pour l’instant. Vous devrez rejoindre le programme [Office Insider](https://insider.office.com/join) pour essayer les fonctionnalités d’aperçu.  Un bon moyen de tester les fonctionnalités en préversion consiste à utiliser un abonnement Office 365. Si vous n’avez pas d'abonnement Office 365, vous pouvez obtenir une version Office 365 gratuite et renouvelable de 90 jours en rejoignant le [Programme pour les développeurs Office 365](https://developer.microsoft.com/office/dev-program).

Si un problème se produit alors que votre fonction personnalisée s'exécute, vous devez retourner une erreur pour informer l’utilisateur. Si vous avez des exigences spécifiques concernant les paramètres, comme des nombres positifs uniquement, vous devez tester les paramètres et générer une erreur s’ils ne sont pas corrects. Vous pouvez également utiliser un bloc `try`-`catch` pour détecter les erreurs qui se produisent pendant que votre fonction personnalisée s’exécute.

## <a name="detect-and-throw-an-error"></a>Détecter et générer une erreur

Examinons un cas où vous devez vous assurer qu’un paramètre de code postal est dans le format correct pour que la fonction personnalisée fonctionne. La fonction personnalisée suivante utilise une expression régulière pour vérifier le code postal. S’il est correct, elle recherche la ville (dans une autre fonction) et retourne la valeur. S’il n’est pas correct, elle retourne une erreur `#VALUE!` à la cellule.

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

L’objet `CustomFunctions.Error` est utilisé pour retourner une erreur à la cellule. Lorsque vous créez l’objet, spécifiez l’erreur que vous voulez utiliser à l’aide de l’une des valeurs enum `ErrorCode` suivantes.


|Valeur enum ErrorCode  |Valeur de la cellule Excel  |Signification  |
|---------------|---------|---------|
|`invalidValue`   | `#VALUE!` | Le type d’une valeur utilisée dans la formule n’est pas bon. |
|`notAvailable`   | `#N/A`    | La fonction ou le service n’est pas disponible. |
|`divisionByZero` | `#DIV/0`  | Sachez que JavaScript autorise la division par zéro, donc vous devez écrire un gestionnaire d’erreurs avec attention pour détecter cette condition. |
|`invalidNumber`  | `#NUM!`   | Un problème s’est produit au niveau du nombre utilisé dans la formule. |
|`nullReference`  | `#NULL!`  | Les plages dans la formule ne se croisent pas. |

L’exemple de code suivant montre comment créer et retourner une erreur pour un nombre non valide (`#NUM!`).

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

Lorsque vous retournez une erreur `#VALUE!`, vous pouvez aussi ajouter un message personnalisé qui apparaîtra dans une fenêtre contextuelle quand l’utilisateur pointera sur la cellule. L’exemple suivant montre comment retourner un message d’erreur personnalisé.

```typescript
// You can only return a custom error message with the #VALUE! error
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a>Utiliser des blocs try-catch

Normalement, vous devez utiliser des blocs `try`-`catch` dans votre fonction personnalisée pour détecter toute erreur potentielle qui se produit. Si vous ne gérez pas les exceptions dans votre code, celles-ci sont retournées à Excel. Par défaut, Excel retourne `#VALUE!` pour une exception non gérée.

Dans l’exemple de code suivant, la fonction personnalisée effectue un appel d’extraction à un service REST. Il est possible que l’appel échoue, par exemple, si le service REST retourne une erreur ou si le réseau est défaillant. Si c’est le cas, la fonction personnalisée retourne `#N/A` pour indiquer que l’appel web a échoué.


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
