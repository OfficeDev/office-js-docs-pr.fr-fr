---
ms.date: 05/06/2020
description: 'Gérer et retourner des erreurs comme #NULL! à partir de votre fonction personnalisée'
title: Gérer et retourner des erreurs à partir de votre fonction personnalisée (préversion)
localization_priority: Normal
ms.openlocfilehash: 6ded6a03151777c30fe5037b373272c04fc64620
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609316"
---
# <a name="handle-and-return-errors-from-your-custom-function-preview"></a><span data-ttu-id="88942-104">Gérer et retourner des erreurs à partir de votre fonction personnalisée (préversion)</span><span class="sxs-lookup"><span data-stu-id="88942-104">Handle and return errors from your custom function (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="88942-105">Les fonctionnalités décrites dans cet article sont actuellement en préversion et peuvent faire l’objet de modifications.</span><span class="sxs-lookup"><span data-stu-id="88942-105">The features described in this article are currently in preview and subject to change.</span></span> <span data-ttu-id="88942-106">Elles ne sont pas prises en charge dans les environnements de production pour l’instant.</span><span class="sxs-lookup"><span data-stu-id="88942-106">They are not currently supported for use in production environments.</span></span> <span data-ttu-id="88942-107">Vous devrez rejoindre le programme [Office Insider](https://insider.office.com/join) pour essayer les fonctionnalités d’aperçu.</span><span class="sxs-lookup"><span data-stu-id="88942-107">You will need to join the [Office Insider](https://insider.office.com/join) program to try the preview features.</span></span>  <span data-ttu-id="88942-108">Un bon moyen de tester les fonctionnalités en préversion consiste à utiliser un abonnement Office 365.</span><span class="sxs-lookup"><span data-stu-id="88942-108">A good way to try out preview features is by using an Office 365 subscription.</span></span> <span data-ttu-id="88942-109">Si vous n’avez pas d'abonnement Office 365, vous pouvez obtenir une version Office 365 gratuite et renouvelable de 90 jours en rejoignant le [Programme pour les développeurs Office 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="88942-109">If you don't already have an Office 365 subscription, you can get a free, 90-day renewable Office 365 subscription by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span>

<span data-ttu-id="88942-110">Si un problème se présente lors de l’exécution de votre fonction personnalisée, renvoyez une erreur pour informer l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="88942-110">If something goes wrong while your custom function runs, return an error to inform the user.</span></span> <span data-ttu-id="88942-111">Si vous avez des exigences de paramètres spécifiques, telles que des nombres positifs, testez les paramètres et générez une erreur s’ils ne sont pas corrects.</span><span class="sxs-lookup"><span data-stu-id="88942-111">If you have specific parameter requirements, such as only positive numbers, test the parameters and throw an error if they aren't correct.</span></span> <span data-ttu-id="88942-112">Vous pouvez également utiliser un bloc `try`-`catch` pour détecter les erreurs qui se produisent pendant que votre fonction personnalisée s’exécute.</span><span class="sxs-lookup"><span data-stu-id="88942-112">You can also use a `try`-`catch` block to catch any errors that occur while your custom function runs.</span></span>

## <a name="detect-and-throw-an-error"></a><span data-ttu-id="88942-113">Détecter et générer une erreur</span><span class="sxs-lookup"><span data-stu-id="88942-113">Detect and throw an error</span></span>

<span data-ttu-id="88942-114">Examinons un cas où vous devez vous assurer qu’un paramètre de code postal est dans le format correct pour que la fonction personnalisée fonctionne.</span><span class="sxs-lookup"><span data-stu-id="88942-114">Let's look at a case where you need to ensure that a zip code parameter is in the correct format for the custom function to work.</span></span> <span data-ttu-id="88942-115">La fonction personnalisée suivante utilise une expression régulière pour vérifier le code postal.</span><span class="sxs-lookup"><span data-stu-id="88942-115">The following custom function uses a regular expression to check the zip code.</span></span> <span data-ttu-id="88942-116">S’il est correct, il recherche la ville à l’aide d’une autre fonction et renvoie la valeur.</span><span class="sxs-lookup"><span data-stu-id="88942-116">If it is correct, then it will look up the city using another function, and return the value.</span></span> <span data-ttu-id="88942-117">Si ce n’est pas le cas, elle renvoie une `#VALUE!` erreur à la cellule.</span><span class="sxs-lookup"><span data-stu-id="88942-117">If it isn't correct, it returns a `#VALUE!` error to the cell.</span></span>

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

## <a name="the-customfunctionserror-object"></a><span data-ttu-id="88942-118">Objet CustomFunctions.Error</span><span class="sxs-lookup"><span data-stu-id="88942-118">The CustomFunctions.Error object</span></span>

<span data-ttu-id="88942-119">L’objet `CustomFunctions.Error` est utilisé pour retourner une erreur à la cellule.</span><span class="sxs-lookup"><span data-stu-id="88942-119">The `CustomFunctions.Error` object is used to return an error back to the cell.</span></span> <span data-ttu-id="88942-120">Lorsque vous créez l’objet, spécifiez l’erreur que vous voulez utiliser à l’aide de l’une des valeurs enum `ErrorCode` suivantes.</span><span class="sxs-lookup"><span data-stu-id="88942-120">When you create the object, specify which error you want to use by using one of the following `ErrorCode` enum values.</span></span>


|<span data-ttu-id="88942-121">Valeur enum ErrorCode</span><span class="sxs-lookup"><span data-stu-id="88942-121">ErrorCode enum value</span></span>  |<span data-ttu-id="88942-122">Valeur de la cellule Excel</span><span class="sxs-lookup"><span data-stu-id="88942-122">Excel cell value</span></span>  |<span data-ttu-id="88942-123">Signification</span><span class="sxs-lookup"><span data-stu-id="88942-123">Meaning</span></span>  |
|---------------|---------|---------|
|`invalidValue`   | `#VALUE!` | <span data-ttu-id="88942-124">Le type d’une valeur utilisée dans la formule n’est pas bon.</span><span class="sxs-lookup"><span data-stu-id="88942-124">A value used in the formula is the wrong type.</span></span> |
|`notAvailable`   | `#N/A`    | <span data-ttu-id="88942-125">La fonction ou le service n’est pas disponible.</span><span class="sxs-lookup"><span data-stu-id="88942-125">The function or service isn't available.</span></span> |
|`divisionByZero` | `#DIV/0`  | <span data-ttu-id="88942-126">Sachez que JavaScript autorise la division par zéro, donc vous devez écrire un gestionnaire d’erreurs avec attention pour détecter cette condition.</span><span class="sxs-lookup"><span data-stu-id="88942-126">Be aware that JavaScript allows division by zero so you need to write an error handler carefully to detect this condition.</span></span> |
|`invalidNumber`  | `#NUM!`   | <span data-ttu-id="88942-127">Un problème s’est produit au niveau du nombre utilisé dans la formule.</span><span class="sxs-lookup"><span data-stu-id="88942-127">There is a problem with the number used in the formula</span></span> |
|`nullReference`  | `#NULL!`  | <span data-ttu-id="88942-128">Les plages de la formule ne se croisent pas.</span><span class="sxs-lookup"><span data-stu-id="88942-128">The ranges in the formula don't intersect.</span></span> |

<span data-ttu-id="88942-129">L’exemple de code suivant montre comment créer et retourner une erreur pour un nombre non valide (`#NUM!`).</span><span class="sxs-lookup"><span data-stu-id="88942-129">The following code sample shows how to create and return an error for an invalid number (`#NUM!`).</span></span>

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

<span data-ttu-id="88942-130">Lorsque vous retournez une erreur `#VALUE!`, vous pouvez aussi ajouter un message personnalisé qui apparaîtra dans une fenêtre contextuelle quand l’utilisateur pointera sur la cellule.</span><span class="sxs-lookup"><span data-stu-id="88942-130">When you return a `#VALUE!` error you can also include a custom message that will be shown in a popup when the user hovers over the cell.</span></span> <span data-ttu-id="88942-131">L’exemple suivant montre comment retourner un message d’erreur personnalisé.</span><span class="sxs-lookup"><span data-stu-id="88942-131">The following example shows how to return a custom error message.</span></span>

```typescript
// You can only return a custom error message with the #VALUE! error
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a><span data-ttu-id="88942-132">Utiliser des blocs try-catch</span><span class="sxs-lookup"><span data-stu-id="88942-132">Use try-catch blocks</span></span>

<span data-ttu-id="88942-133">En règle générale, utilisez des `try` - `catch` blocs dans votre fonction personnalisée pour intercepter les erreurs potentielles qui se produisent.</span><span class="sxs-lookup"><span data-stu-id="88942-133">In general, use `try`-`catch` blocks in your custom function to catch any potential errors that occur.</span></span> <span data-ttu-id="88942-134">Si vous ne gérez pas les exceptions dans votre code, celles-ci sont retournées à Excel.</span><span class="sxs-lookup"><span data-stu-id="88942-134">If you do not handle exceptions in your code, they will be returned to Excel.</span></span> <span data-ttu-id="88942-135">Par défaut, Excel retourne `#VALUE!` pour une exception non gérée.</span><span class="sxs-lookup"><span data-stu-id="88942-135">By default, Excel returns `#VALUE!` for an unhandled exception.</span></span>

<span data-ttu-id="88942-136">Dans l’exemple de code suivant, la fonction personnalisée effectue un appel d’extraction à un service REST.</span><span class="sxs-lookup"><span data-stu-id="88942-136">In the following code sample, the custom function makes a fetch call to a REST service.</span></span> <span data-ttu-id="88942-137">Il est possible que l’appel échoue, par exemple, si le service REST retourne une erreur ou si le réseau est défaillant.</span><span class="sxs-lookup"><span data-stu-id="88942-137">It's possible that the call will fail, for example, if the REST service returns an error or the network goes down.</span></span> <span data-ttu-id="88942-138">Si c’est le cas, la fonction personnalisée retourne `#N/A` pour indiquer que l’appel web a échoué.</span><span class="sxs-lookup"><span data-stu-id="88942-138">If this happens, the custom function will return `#N/A` to indicate the web call failed.</span></span>


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

## <a name="next-steps"></a><span data-ttu-id="88942-139">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="88942-139">Next steps</span></span>

<span data-ttu-id="88942-140">Découvrez comment [résoudre les problèmes liés à vos fonctions personnalisées](custom-functions-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="88942-140">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="88942-141">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="88942-141">See also</span></span>

* [<span data-ttu-id="88942-142">Débogage des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="88942-142">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="88942-143">Configuration requise de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="88942-143">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="88942-144">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="88942-144">Create custom functions in Excel</span></span>](custom-functions-overview.md)
