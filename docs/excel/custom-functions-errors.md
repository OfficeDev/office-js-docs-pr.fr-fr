---
ms.date: 09/15/2020
description: 'Gérer et retourner des erreurs comme #NULL! à partir de votre fonction personnalisée.'
title: Gérer et renvoyer des erreurs à partir de votre fonction personnalisée
localization_priority: Normal
ms.openlocfilehash: 5da68417aa52f1d14340c8c8a46f4943ffd2d223
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819531"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a><span data-ttu-id="c58b6-104">Gérer et renvoyer des erreurs à partir de votre fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="c58b6-104">Handle and return errors from your custom function</span></span>

> [!NOTE]
> <span data-ttu-id="c58b6-105">Les fonctionnalités décrites dans cet article sont actuellement en préversion et peuvent faire l’objet de modifications.</span><span class="sxs-lookup"><span data-stu-id="c58b6-105">The features described in this article are currently in preview and subject to change.</span></span> <span data-ttu-id="c58b6-106">Elles ne sont pas prises en charge dans les environnements de production pour l’instant.</span><span class="sxs-lookup"><span data-stu-id="c58b6-106">They are not currently supported for use in production environments.</span></span> <span data-ttu-id="c58b6-107">Vous devrez rejoindre le programme [Office Insider](https://insider.office.com/join) pour essayer les fonctionnalités d’aperçu.</span><span class="sxs-lookup"><span data-stu-id="c58b6-107">You will need to join the [Office Insider](https://insider.office.com/join) program to try the preview features.</span></span>  <span data-ttu-id="c58b6-108">Pour tester les fonctionnalités d’aperçu, il est recommandé d’utiliser un abonnement Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="c58b6-108">A good way to try out preview features is by using a Microsoft 365 subscription.</span></span> <span data-ttu-id="c58b6-109">Si vous ne disposez pas déjà d’un abonnement Microsoft 365, vous pouvez obtenir gratuitement un abonnement Microsoft 365 renouvelable 90 jours en joignant le [programme de développement microsoft 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="c58b6-109">If you don't already have a Microsoft 365 subscription, you can get a free, 90-day renewable Microsoft 365 subscription by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

<span data-ttu-id="c58b6-110">Si un problème se présente lors de l’exécution de votre fonction personnalisée, renvoyez une erreur pour informer l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c58b6-110">If something goes wrong while your custom function runs, return an error to inform the user.</span></span> <span data-ttu-id="c58b6-111">Si vous avez des exigences de paramètres spécifiques, telles que des nombres positifs, testez les paramètres et générez une erreur s’ils ne sont pas corrects.</span><span class="sxs-lookup"><span data-stu-id="c58b6-111">If you have specific parameter requirements, such as only positive numbers, test the parameters and throw an error if they aren't correct.</span></span> <span data-ttu-id="c58b6-112">Vous pouvez également utiliser un bloc `try`-`catch` pour détecter les erreurs qui se produisent pendant que votre fonction personnalisée s’exécute.</span><span class="sxs-lookup"><span data-stu-id="c58b6-112">You can also use a `try`-`catch` block to catch any errors that occur while your custom function runs.</span></span>

## <a name="detect-and-throw-an-error"></a><span data-ttu-id="c58b6-113">Détecter et générer une erreur</span><span class="sxs-lookup"><span data-stu-id="c58b6-113">Detect and throw an error</span></span>

<span data-ttu-id="c58b6-114">Examinons un cas où vous devez vous assurer qu’un paramètre de code postal est dans le format correct pour que la fonction personnalisée fonctionne.</span><span class="sxs-lookup"><span data-stu-id="c58b6-114">Let's look at a case where you need to ensure that a zip code parameter is in the correct format for the custom function to work.</span></span> <span data-ttu-id="c58b6-115">La fonction personnalisée suivante utilise une expression régulière pour vérifier le code postal.</span><span class="sxs-lookup"><span data-stu-id="c58b6-115">The following custom function uses a regular expression to check the zip code.</span></span> <span data-ttu-id="c58b6-116">S’il est correct, il recherche la ville à l’aide d’une autre fonction et renvoie la valeur.</span><span class="sxs-lookup"><span data-stu-id="c58b6-116">If it is correct, then it will look up the city using another function, and return the value.</span></span> <span data-ttu-id="c58b6-117">Si ce n’est pas le cas, elle renvoie une `#VALUE!` erreur à la cellule.</span><span class="sxs-lookup"><span data-stu-id="c58b6-117">If it isn't correct, it returns a `#VALUE!` error to the cell.</span></span>

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

## <a name="the-customfunctionserror-object"></a><span data-ttu-id="c58b6-118">Objet CustomFunctions.Error</span><span class="sxs-lookup"><span data-stu-id="c58b6-118">The CustomFunctions.Error object</span></span>

<span data-ttu-id="c58b6-119">L’objet `CustomFunctions.Error` est utilisé pour retourner une erreur à la cellule.</span><span class="sxs-lookup"><span data-stu-id="c58b6-119">The `CustomFunctions.Error` object is used to return an error back to the cell.</span></span> <span data-ttu-id="c58b6-120">Lorsque vous créez l’objet, spécifiez l’erreur que vous voulez utiliser à l’aide de l’une des valeurs enum `ErrorCode` suivantes.</span><span class="sxs-lookup"><span data-stu-id="c58b6-120">When you create the object, specify which error you want to use by using one of the following `ErrorCode` enum values.</span></span>


|<span data-ttu-id="c58b6-121">Valeur enum ErrorCode</span><span class="sxs-lookup"><span data-stu-id="c58b6-121">ErrorCode enum value</span></span>  |<span data-ttu-id="c58b6-122">Valeur de la cellule Excel</span><span class="sxs-lookup"><span data-stu-id="c58b6-122">Excel cell value</span></span>  |<span data-ttu-id="c58b6-123">Signification</span><span class="sxs-lookup"><span data-stu-id="c58b6-123">Meaning</span></span>  |
|---------------|---------|---------|
|`invalidValue`   | `#VALUE!` | <span data-ttu-id="c58b6-124">Le type d’une valeur utilisée dans la formule n’est pas bon.</span><span class="sxs-lookup"><span data-stu-id="c58b6-124">A value used in the formula is the wrong type.</span></span> |
|`notAvailable`   | `#N/A`    | <span data-ttu-id="c58b6-125">La fonction ou le service n’est pas disponible.</span><span class="sxs-lookup"><span data-stu-id="c58b6-125">The function or service isn't available.</span></span> |
|`divisionByZero` | `#DIV/0`  | <span data-ttu-id="c58b6-126">Sachez que JavaScript autorise la division par zéro, donc vous devez écrire un gestionnaire d’erreurs avec attention pour détecter cette condition.</span><span class="sxs-lookup"><span data-stu-id="c58b6-126">Be aware that JavaScript allows division by zero so you need to write an error handler carefully to detect this condition.</span></span> |
|`invalidNumber`  | `#NUM!`   | <span data-ttu-id="c58b6-127">Un problème s’est produit au niveau du nombre utilisé dans la formule.</span><span class="sxs-lookup"><span data-stu-id="c58b6-127">There is a problem with the number used in the formula</span></span> |
|`nullReference`  | `#NULL!`  | <span data-ttu-id="c58b6-128">Les plages de la formule ne se croisent pas.</span><span class="sxs-lookup"><span data-stu-id="c58b6-128">The ranges in the formula don't intersect.</span></span> |

<span data-ttu-id="c58b6-129">L’exemple de code suivant montre comment créer et retourner une erreur pour un nombre non valide (`#NUM!`).</span><span class="sxs-lookup"><span data-stu-id="c58b6-129">The following code sample shows how to create and return an error for an invalid number (`#NUM!`).</span></span>

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

<span data-ttu-id="c58b6-130">Lorsque vous retournez une erreur `#VALUE!`, vous pouvez aussi ajouter un message personnalisé qui apparaîtra dans une fenêtre contextuelle quand l’utilisateur pointera sur la cellule.</span><span class="sxs-lookup"><span data-stu-id="c58b6-130">When you return a `#VALUE!` error you can also include a custom message that will be shown in a popup when the user hovers over the cell.</span></span> <span data-ttu-id="c58b6-131">L’exemple suivant montre comment retourner un message d’erreur personnalisé.</span><span class="sxs-lookup"><span data-stu-id="c58b6-131">The following example shows how to return a custom error message.</span></span>

```typescript
// You can only return a custom error message with the #VALUE! error
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a><span data-ttu-id="c58b6-132">Utiliser des blocs try-catch</span><span class="sxs-lookup"><span data-stu-id="c58b6-132">Use try-catch blocks</span></span>

<span data-ttu-id="c58b6-133">En règle générale, utilisez des `try` - `catch` blocs dans votre fonction personnalisée pour intercepter les erreurs potentielles qui se produisent.</span><span class="sxs-lookup"><span data-stu-id="c58b6-133">In general, use `try`-`catch` blocks in your custom function to catch any potential errors that occur.</span></span> <span data-ttu-id="c58b6-134">Si vous ne gérez pas les exceptions dans votre code, celles-ci sont retournées à Excel.</span><span class="sxs-lookup"><span data-stu-id="c58b6-134">If you do not handle exceptions in your code, they will be returned to Excel.</span></span> <span data-ttu-id="c58b6-135">Par défaut, Excel retourne `#VALUE!` pour une exception non gérée.</span><span class="sxs-lookup"><span data-stu-id="c58b6-135">By default, Excel returns `#VALUE!` for an unhandled exception.</span></span>

<span data-ttu-id="c58b6-136">Dans l’exemple de code suivant, la fonction personnalisée effectue un appel d’extraction à un service REST.</span><span class="sxs-lookup"><span data-stu-id="c58b6-136">In the following code sample, the custom function makes a fetch call to a REST service.</span></span> <span data-ttu-id="c58b6-137">Il est possible que l’appel échoue, par exemple, si le service REST retourne une erreur ou si le réseau est défaillant.</span><span class="sxs-lookup"><span data-stu-id="c58b6-137">It's possible that the call will fail, for example, if the REST service returns an error or the network goes down.</span></span> <span data-ttu-id="c58b6-138">Si c’est le cas, la fonction personnalisée retourne `#N/A` pour indiquer que l’appel web a échoué.</span><span class="sxs-lookup"><span data-stu-id="c58b6-138">If this happens, the custom function will return `#N/A` to indicate the web call failed.</span></span>


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

## <a name="next-steps"></a><span data-ttu-id="c58b6-139">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="c58b6-139">Next steps</span></span>

<span data-ttu-id="c58b6-140">Découvrez comment [résoudre les problèmes liés à vos fonctions personnalisées](custom-functions-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="c58b6-140">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="c58b6-141">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c58b6-141">See also</span></span>

* [<span data-ttu-id="c58b6-142">Débogage des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="c58b6-142">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="c58b6-143">Configuration requise de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="c58b6-143">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="c58b6-144">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="c58b6-144">Create custom functions in Excel</span></span>](custom-functions-overview.md)
