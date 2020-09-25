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
# <a name="handle-and-return-errors-from-your-custom-function"></a><span data-ttu-id="e77cc-104">Gérer et renvoyer des erreurs à partir de votre fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="e77cc-104">Handle and return errors from your custom function</span></span>

<span data-ttu-id="e77cc-105">Si un problème se présente lors de l’exécution de votre fonction personnalisée, renvoyez une erreur pour informer l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e77cc-105">If something goes wrong while your custom function runs, return an error to inform the user.</span></span> <span data-ttu-id="e77cc-106">Si vous avez des exigences de paramètres spécifiques, telles que des nombres positifs, testez les paramètres et générez une erreur s’ils ne sont pas corrects.</span><span class="sxs-lookup"><span data-stu-id="e77cc-106">If you have specific parameter requirements, such as only positive numbers, test the parameters and throw an error if they aren't correct.</span></span> <span data-ttu-id="e77cc-107">Vous pouvez également utiliser un bloc `try`-`catch` pour détecter les erreurs qui se produisent pendant que votre fonction personnalisée s’exécute.</span><span class="sxs-lookup"><span data-stu-id="e77cc-107">You can also use a `try`-`catch` block to catch any errors that occur while your custom function runs.</span></span>

## <a name="detect-and-throw-an-error"></a><span data-ttu-id="e77cc-108">Détecter et générer une erreur</span><span class="sxs-lookup"><span data-stu-id="e77cc-108">Detect and throw an error</span></span>

<span data-ttu-id="e77cc-109">Examinons un cas où vous devez vous assurer qu’un paramètre de code postal est dans le format correct pour que la fonction personnalisée fonctionne.</span><span class="sxs-lookup"><span data-stu-id="e77cc-109">Let's look at a case where you need to ensure that a zip code parameter is in the correct format for the custom function to work.</span></span> <span data-ttu-id="e77cc-110">La fonction personnalisée suivante utilise une expression régulière pour vérifier le code postal.</span><span class="sxs-lookup"><span data-stu-id="e77cc-110">The following custom function uses a regular expression to check the zip code.</span></span> <span data-ttu-id="e77cc-111">Si le format de code postal est correct, il recherche la ville à l’aide d’une autre fonction et renvoie la valeur.</span><span class="sxs-lookup"><span data-stu-id="e77cc-111">If the zip code format is correct, then it will look up the city using another function and return the value.</span></span> <span data-ttu-id="e77cc-112">Si le format n’est pas valide, la fonction renvoie une `#VALUE!` erreur à la cellule.</span><span class="sxs-lookup"><span data-stu-id="e77cc-112">If the format isn't valid, the function returns a `#VALUE!` error to the cell.</span></span>

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

## <a name="the-customfunctionserror-object"></a><span data-ttu-id="e77cc-113">Objet CustomFunctions.Error</span><span class="sxs-lookup"><span data-stu-id="e77cc-113">The CustomFunctions.Error object</span></span>

<span data-ttu-id="e77cc-114">L’objet [CustomFunctions. Error](/javascript/api/custom-functions-runtime/customfunctions.error) est utilisé pour renvoyer une erreur à la cellule.</span><span class="sxs-lookup"><span data-stu-id="e77cc-114">The [CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) object is used to return an error back to the cell.</span></span> <span data-ttu-id="e77cc-115">Lorsque vous créez l’objet, spécifiez l’erreur que vous souhaitez utiliser en choisissant l’une des `ErrorCode` valeurs d’énumération suivantes.</span><span class="sxs-lookup"><span data-stu-id="e77cc-115">When you create the object, specify which error you want to use by choosing one of the following `ErrorCode` enum values.</span></span>


|<span data-ttu-id="e77cc-116">Valeur enum ErrorCode</span><span class="sxs-lookup"><span data-stu-id="e77cc-116">ErrorCode enum value</span></span>  |<span data-ttu-id="e77cc-117">Valeur de la cellule Excel</span><span class="sxs-lookup"><span data-stu-id="e77cc-117">Excel cell value</span></span>  |<span data-ttu-id="e77cc-118">Description</span><span class="sxs-lookup"><span data-stu-id="e77cc-118">Description</span></span>  |
|---------------|---------|---------|
|`divisionByZero` | `#DIV/0`  | <span data-ttu-id="e77cc-119">La fonction tente d’effectuer une division par zéro.</span><span class="sxs-lookup"><span data-stu-id="e77cc-119">The function is attempting to divide by zero.</span></span> |
|`invalidName`    | `#NAME?`  | <span data-ttu-id="e77cc-120">Il y a une faute de frappe dans le nom de la fonction.</span><span class="sxs-lookup"><span data-stu-id="e77cc-120">There is a typo in the function name.</span></span> <span data-ttu-id="e77cc-121">Notez que cette erreur est prise en charge en tant qu’erreur d’entrée d’une fonction personnalisée, mais pas en tant qu’erreur de sortie d’une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="e77cc-121">Note that this error is supported as a custom function input error, but not as a custom function output error.</span></span> | 
|`invalidNumber`  | `#NUM!`   | <span data-ttu-id="e77cc-122">Il y a un problème avec un nombre dans la formule.</span><span class="sxs-lookup"><span data-stu-id="e77cc-122">There is a problem with a number in the formula.</span></span> |
|`invalidReference` | `#REF!` | <span data-ttu-id="e77cc-123">La fonction fait référence à une cellule non valide.</span><span class="sxs-lookup"><span data-stu-id="e77cc-123">The function refers to an invalid cell.</span></span> <span data-ttu-id="e77cc-124">Notez que cette erreur est prise en charge en tant qu’erreur d’entrée d’une fonction personnalisée, mais pas en tant qu’erreur de sortie d’une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="e77cc-124">Note that this error is supported as a custom function input error, but not as a custom function output error.</span></span>|
|`invalidValue`   | `#VALUE!` | <span data-ttu-id="e77cc-125">La valeur de la formule est de type incorrect.</span><span class="sxs-lookup"><span data-stu-id="e77cc-125">A value in the formula is of the wrong type.</span></span> |
|`notAvailable`   | `#N/A`    | <span data-ttu-id="e77cc-126">La fonction ou le service n’est pas disponible.</span><span class="sxs-lookup"><span data-stu-id="e77cc-126">The function or service isn't available.</span></span> |
|`nullReference`  | `#NULL!`  | <span data-ttu-id="e77cc-127">Les plages de la formule ne se croisent pas.</span><span class="sxs-lookup"><span data-stu-id="e77cc-127">The ranges in the formula don't intersect.</span></span> |

<span data-ttu-id="e77cc-128">L’exemple de code suivant montre comment créer et retourner une erreur pour un nombre non valide (`#NUM!`).</span><span class="sxs-lookup"><span data-stu-id="e77cc-128">The following code sample shows how to create and return an error for an invalid number (`#NUM!`).</span></span>

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

<span data-ttu-id="e77cc-129">Les `#VALUE!` `#N/A` Erreurs et prennent également en charge les messages d’erreur personnalisés.</span><span class="sxs-lookup"><span data-stu-id="e77cc-129">The `#VALUE!` and `#N/A` errors also support custom error messages.</span></span> <span data-ttu-id="e77cc-130">Les messages d’erreur personnalisés s’affichent dans le menu indicateur d’erreur, accessible en plaçant le curseur sur l’indicateur d’erreur sur chaque cellule avec une erreur.</span><span class="sxs-lookup"><span data-stu-id="e77cc-130">Custom error messages are displayed in the error indicator menu, which is accessed by hovering over the error flag on each cell with an error.</span></span> <span data-ttu-id="e77cc-131">L’exemple suivant montre comment renvoyer un message d’erreur personnalisé avec l' `#VALUE!` erreur.</span><span class="sxs-lookup"><span data-stu-id="e77cc-131">The following example shows how to return a custom error message with the `#VALUE!` error.</span></span>

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a><span data-ttu-id="e77cc-132">Utiliser des blocs try-catch</span><span class="sxs-lookup"><span data-stu-id="e77cc-132">Use try-catch blocks</span></span>

<span data-ttu-id="e77cc-133">En règle générale, utilisez des `try` - `catch` blocs dans votre fonction personnalisée pour intercepter les erreurs potentielles qui se produisent.</span><span class="sxs-lookup"><span data-stu-id="e77cc-133">In general, use `try`-`catch` blocks in your custom function to catch any potential errors that occur.</span></span> <span data-ttu-id="e77cc-134">Si vous ne gérez pas les exceptions dans votre code, celles-ci sont retournées à Excel.</span><span class="sxs-lookup"><span data-stu-id="e77cc-134">If you do not handle exceptions in your code, they will be returned to Excel.</span></span> <span data-ttu-id="e77cc-135">Par défaut, Excel renvoie `#VALUE!` des exceptions ou des erreurs non gérées.</span><span class="sxs-lookup"><span data-stu-id="e77cc-135">By default, Excel returns `#VALUE!` for unhandled errors or exceptions.</span></span>

<span data-ttu-id="e77cc-136">Dans l’exemple de code suivant, la fonction personnalisée effectue un appel d’extraction à un service REST.</span><span class="sxs-lookup"><span data-stu-id="e77cc-136">In the following code sample, the custom function makes a fetch call to a REST service.</span></span> <span data-ttu-id="e77cc-137">Il est possible que l’appel échoue, par exemple, si le service REST retourne une erreur ou si le réseau est défaillant.</span><span class="sxs-lookup"><span data-stu-id="e77cc-137">It's possible that the call will fail, for example, if the REST service returns an error or the network goes down.</span></span> <span data-ttu-id="e77cc-138">Dans ce cas, la fonction personnalisée renvoie `#N/A` pour indiquer que l’appel Web a échoué.</span><span class="sxs-lookup"><span data-stu-id="e77cc-138">If this happens, the custom function will return `#N/A` to indicate that the web call failed.</span></span>


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

## <a name="next-steps"></a><span data-ttu-id="e77cc-139">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="e77cc-139">Next steps</span></span>

<span data-ttu-id="e77cc-140">Découvrez comment [résoudre les problèmes liés à vos fonctions personnalisées](custom-functions-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="e77cc-140">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="e77cc-141">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e77cc-141">See also</span></span>

* [<span data-ttu-id="e77cc-142">Débogage des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="e77cc-142">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="e77cc-143">Configuration requise de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="e77cc-143">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="e77cc-144">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="e77cc-144">Create custom functions in Excel</span></span>](custom-functions-overview.md)
