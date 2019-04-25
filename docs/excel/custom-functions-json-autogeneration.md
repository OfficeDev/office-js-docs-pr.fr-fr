---
ms.date: 04/03/2019
description: Utiliser les balises JSDOC pour créer dynamiquement vos fonctions personnalisées de métadonnées JSON.
title: Créer les métadonnées JSON pour des fonctions personnalisées (aperçu)
localization_priority: Priority
ms.openlocfilehash: 2efe2a9a5a83ba60ef327273d5bd599f82916d48
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449252"
---
# <a name="create-json-metadata-for-custom-functions-preview"></a><span data-ttu-id="8bd8f-103">Créer les métadonnées JSON pour des fonctions personnalisées (aperçu)</span><span class="sxs-lookup"><span data-stu-id="8bd8f-103">Create JSON metadata for custom functions (preview)</span></span>

<span data-ttu-id="8bd8f-104">Si vous écrivez une fonction Excel personnalisée en JavaScript ou TypeScript, vous pouvez utiliser les balises JSDoc pour la détailler en ajoutant des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="8bd8f-105">Les balises JSDoc sont ensuite utilisées lors de la génération pour créer le [fichier de métadonnées JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="8bd8f-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="8bd8f-106">En utilisant des balises JSDoc, vous n’avez plus besoin de modifier manuellement le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

<span data-ttu-id="8bd8f-107">Ajoutez la balise `@customfunction` dans les commentaires du code d’une fonction JavaScript ou TypeScript pour indiquer qu’il s’agit d’une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="8bd8f-108">Vous pouvez fournir les types de paramètres de la fonction en utilisant la balise[@param](#param)dans JavaScript, ou en précisant le [type de fonction](https://www.typescriptlang.org/docs/handbook/functions.html) dans TypeScript.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="8bd8f-109">Si vous voulez en savoir plus, veuillez consulter les sections relatives à la balise[@param](#param) et aux sections[types](#types).</span><span class="sxs-lookup"><span data-stu-id="8bd8f-109">For more information, see the [@param](#param) tag and [Types](#types) section.</span></span>

## <a name="jsdoc-tags"></a><span data-ttu-id="8bd8f-110">Balises JSDoc</span><span class="sxs-lookup"><span data-stu-id="8bd8f-110">JSDoc Tags</span></span>
<span data-ttu-id="8bd8f-111">Voici quelles sont les balises JSDoc prises en charge dans les fonctions Excel personnalisées :</span><span class="sxs-lookup"><span data-stu-id="8bd8f-111">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [<span data-ttu-id="8bd8f-112">@ annulable</span><span class="sxs-lookup"><span data-stu-id="8bd8f-112">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="8bd8f-113">[@fonctionpersonnalisée](#customfunction)nom id</span><span class="sxs-lookup"><span data-stu-id="8bd8f-113">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="8bd8f-114">url[@urlaide](#helpurl)</span><span class="sxs-lookup"><span data-stu-id="8bd8f-114">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="8bd8f-115">[@param](#param) _{type}_ description nom</span><span class="sxs-lookup"><span data-stu-id="8bd8f-115">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="8bd8f-116">@requièreuneadresse</span><span class="sxs-lookup"><span data-stu-id="8bd8f-116">@requiresAddress</span></span>](#requiresAddress)
* <span data-ttu-id="8bd8f-117">[@renvoie](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="8bd8f-117">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="8bd8f-118">@diffusionencontinu</span><span class="sxs-lookup"><span data-stu-id="8bd8f-118">@streaming</span></span>](#streaming)
* [<span data-ttu-id="8bd8f-119">@volatile</span><span class="sxs-lookup"><span data-stu-id="8bd8f-119">@volatile</span></span>](#volatile)

---
### <a name="cancelable"></a><span data-ttu-id="8bd8f-120">@ annulable</span><span class="sxs-lookup"><span data-stu-id="8bd8f-120">@cancelable</span></span>
<a id="cancelable"/>

<span data-ttu-id="8bd8f-121">Indique qu’une fonction personnalisée souhaite effectuer une action lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-121">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="8bd8f-122">Le dernier paramètre de la fonction doit être de type `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-122">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="8bd8f-123">La fonction peut attribuer une fonction à la propriété `oncanceled` pour désigner l’action à effectuer lors de l’annulation de la fonction.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-123">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="8bd8f-124">Si le dernier paramètre de fonction est de type `CustomFunctions.CancelableInvocation`, il sera considéré comme `@cancelable`, même si la balise n’apparaît pas.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-124">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="8bd8f-125">Une fonction ne peut pas contenir les deux balises `@cancelable` et `@streaming`.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-125">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a><span data-ttu-id="8bd8f-126">@fonctionpersonnalisée</span><span class="sxs-lookup"><span data-stu-id="8bd8f-126">@customfunction</span></span>
<a id="customfunction"/>

<span data-ttu-id="8bd8f-127">Syntaxe: @fonctionpersonnalisée_id_ _nom_</span><span class="sxs-lookup"><span data-stu-id="8bd8f-127">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="8bd8f-128">Spécifiez cette balise pour traiter la fonction JavaScript/TypeScript comme une fonction Excel personnalisée.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-128">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="8bd8f-129">Cette balise est requise pour créer des métadonnées pour la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-129">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="8bd8f-130">Vous devez également insérer un appel vers`CustomFunctions.associate("id", functionName);`</span><span class="sxs-lookup"><span data-stu-id="8bd8f-130">There should also be a call to `CustomFunctions.associate("id", functionName);`</span></span>

#### <a name="id"></a><span data-ttu-id="8bd8f-131">id</span><span class="sxs-lookup"><span data-stu-id="8bd8f-131">id</span></span> 

<span data-ttu-id="8bd8f-132">L’id est utilisé en tant qu’identificateur invariant pour la fonction personnalisée stockée dans le document.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-132">The id is used as the invariant identifier for the custom function stored in the document.</span></span> <span data-ttu-id="8bd8f-133">Elle ne doit pas changer.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-133">It should not change.</span></span>

* <span data-ttu-id="8bd8f-134">Si l’id n’est pas fourni, le nom de la fonction JavaScript/TypeScript est convertie en majuscules, et les caractères rejetés sont supprimés.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-134">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="8bd8f-135">L’id doit être unique pour toutes les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-135">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="8bd8f-136">Seuls les caractères alphanumériques majuscules et minuscules (A-Z, a-z, 0-9) et le point (.) sont autorisés.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-136">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

#### <a name="name"></a><span data-ttu-id="8bd8f-137">name</span><span class="sxs-lookup"><span data-stu-id="8bd8f-137">name</span></span>

<span data-ttu-id="8bd8f-138">Fournit le nom d’affichage de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-138">Provides the display name for the custom function.</span></span> 

* <span data-ttu-id="8bd8f-139">Si aucun nom n’est fourni, l’id servira aussi de nom.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-139">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="8bd8f-140">Caractères autorisés : [caractères alphanumériques Unicode](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic) (lettres, chiffres), point (.) et trait de soulignement (\_).</span><span class="sxs-lookup"><span data-stu-id="8bd8f-140">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="8bd8f-141">Doit commencer par une lettre.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-141">Must start with a letter.</span></span>
* <span data-ttu-id="8bd8f-142">Sa longueur maximale est limitée à 128 caractères.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-142">Maximum length is 128 characters.</span></span>

---
### <a name="helpurl"></a><span data-ttu-id="8bd8f-143">@urlaide</span><span class="sxs-lookup"><span data-stu-id="8bd8f-143">@helpurl</span></span>
<a id="helpurl"/>

<span data-ttu-id="8bd8f-144">Syntaxe: @urlaide_url_</span><span class="sxs-lookup"><span data-stu-id="8bd8f-144">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="8bd8f-145">L’_url_ fournie est affichée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-145">The provided _url_ is displayed in Excel.</span></span>

---
### <a name="param"></a><span data-ttu-id="8bd8f-146">@param</span><span class="sxs-lookup"><span data-stu-id="8bd8f-146">@param</span></span>
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="8bd8f-147">JavaScript</span><span class="sxs-lookup"><span data-stu-id="8bd8f-147">JavaScript</span></span>

<span data-ttu-id="8bd8f-148">Syntaxe JavaScript : @param {type} nom_description_</span><span class="sxs-lookup"><span data-stu-id="8bd8f-148">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="8bd8f-149">`{type}`doit spécifier les informations de type entre deux accolades.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-149">`{type}` should specify the type info within curly braces.</span></span> <span data-ttu-id="8bd8f-150">Consultez la section [Types](##types) pour savoir quels types peuvent être utilisés.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-150">See the [Types](##types) for more information about the types which may be used.</span></span> <span data-ttu-id="8bd8f-151">Facultatif : si aucun serveur n’est spécifié, le type `any` sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-151">Optional: if not specified, the type `any` will be used.</span></span>
* <span data-ttu-id="8bd8f-152">`name`spécifie le paramètre auquel s’applique la balise.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-152">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="8bd8f-153">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-153">Required.</span></span>
* <span data-ttu-id="8bd8f-154">`description`fournit la description qui s’affiche dans Excel pour le paramètre de la fonction.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-154">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="8bd8f-155">Facultatif.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-155">Optional.</span></span>

<span data-ttu-id="8bd8f-156">Pour désigner un paramètre de fonction personnalisée comme étant facultatif :</span><span class="sxs-lookup"><span data-stu-id="8bd8f-156">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="8bd8f-157">Placez les crochets autour du nom du paramètre.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-157">Put square brackets around the parameter name.</span></span> <span data-ttu-id="8bd8f-158">Par exemple : `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-158">For example: `@param {string} [text] Optional text`.</span></span>

#### <a name="typescript"></a><span data-ttu-id="8bd8f-159">TypeScript</span><span class="sxs-lookup"><span data-stu-id="8bd8f-159">TypeScript</span></span>

<span data-ttu-id="8bd8f-160">Syntaxe TypeScript : nom @param_description_</span><span class="sxs-lookup"><span data-stu-id="8bd8f-160">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="8bd8f-161">`name`spécifie le paramètre auquel s’applique la balise.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-161">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="8bd8f-162">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-162">Required.</span></span>
* <span data-ttu-id="8bd8f-163">`description`fournit la description qui s’affiche dans Excel pour le paramètre de la fonction.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-163">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="8bd8f-164">Facultatif.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-164">Optional.</span></span>

<span data-ttu-id="8bd8f-165">Consultez la section [Types](##types) pour savoir quels types de paramètres de fonction peuvent être utilisés.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-165">See the [Types](##types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="8bd8f-166">Pour désigner un paramètre de fonction personnalisée comme étant facultatif, effectuez l’une des actions suivantes :</span><span class="sxs-lookup"><span data-stu-id="8bd8f-166">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="8bd8f-167">Utilisez un paramètre facultatif.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-167">Use an optional parameter.</span></span> <span data-ttu-id="8bd8f-168">Par exemple : `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="8bd8f-168">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="8bd8f-169">Définissez ce paramètre sur une valeur par défaut.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-169">Give the parameter a default value.</span></span> <span data-ttu-id="8bd8f-170">Par exemple : `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="8bd8f-170">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="8bd8f-171">Pour consulter une description détaillée du @param, reportez-vous à la page suivante : [JSDoc](http://usejsdoc.org/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="8bd8f-171">For detailed description of the @param see: [JSDoc](http://usejsdoc.org/tags-param.html)</span></span>

---
### <a name="requiresaddress"></a><span data-ttu-id="8bd8f-172">@requièreuneadresse</span><span class="sxs-lookup"><span data-stu-id="8bd8f-172">@requiresAddress</span></span>
<a id="requiresAddress"/>

<span data-ttu-id="8bd8f-173">Indique que l’adresse de la cellule dans laquelle la fonction est évaluée doit être fournie.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-173">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span> 

<span data-ttu-id="8bd8f-174">Le dernier paramètre de la fonction doit être de type `CustomFunctions.Invocation` ou un type dérivé.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-174">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="8bd8f-175">Lorsque la fonction est appelée, la propriété `address` contiendra l’adresse.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-175">When the function is called, the `address` property will contain the address.</span></span>

---
### <a name="returns"></a><span data-ttu-id="8bd8f-176">@renvoie :</span><span class="sxs-lookup"><span data-stu-id="8bd8f-176">@returns</span></span>
<a id="returns"/>

<span data-ttu-id="8bd8f-177">Syntaxe: @renvoie {_type_}</span><span class="sxs-lookup"><span data-stu-id="8bd8f-177">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="8bd8f-178">Fournit le type pour la valeur renvoyée.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-178">Provides the type for the return value.</span></span>

<span data-ttu-id="8bd8f-179">Si `{type}` est omis, les informations de type TypeScript seront utilisées.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-179">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="8bd8f-180">S’il n’existe aucune information définissant le type, ce dernier sera `any`.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-180">If there is no type info, the type will be `any`.</span></span>

---
### <a name="streaming"></a><span data-ttu-id="8bd8f-181">@diffusionencontinu</span><span class="sxs-lookup"><span data-stu-id="8bd8f-181">@streaming</span></span>
<a id="streaming"/>

<span data-ttu-id="8bd8f-182">Utilisé pour indiquer qu’une fonction personnalisée est une fonction diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-182">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="8bd8f-183">Le dernier paramètre doit être de type `CustomFunctions.StreamingInvocation<ResultType>`.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-183">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="8bd8f-184">La fonction doit renvoyer `void`.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-184">The function should return `void`.</span></span>

<span data-ttu-id="8bd8f-185">Les fonctions de diffusion en continu ne renvoient pas de valeurs directement, mais doivent plutôt appeler `setResult(result: ResultType)` en utilisant le dernier paramètre.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-185">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="8bd8f-186">Les exceptions levées par une fonction en continu sont ignorées.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-186">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="8bd8f-187">`setResult()`peut être appelée avec Error pour indiquer un résultat erroné.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-187">`setResult()` may be called with Error to indicate an error result.</span></span>

<span data-ttu-id="8bd8f-188">Vous ne pouvez pas utiliser les balises en diffusion en continu comme [@volatile](#volatile).</span><span class="sxs-lookup"><span data-stu-id="8bd8f-188">Streaming functions cannot be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a><span data-ttu-id="8bd8f-189">@volatile</span><span class="sxs-lookup"><span data-stu-id="8bd8f-189">@volatile</span></span>
<a id="volatile"/>

<span data-ttu-id="8bd8f-190">Une fonction volatile est une fonction dont le résultat peut changer d’un moment à l’autre, même si elle ne récupère pas d’argument ou si ses arguments ne changent pas.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-190">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="8bd8f-191">À chaque calcul, Excel réévalue les cellules contenant des fonctions volatiles, ainsi que toutes leurs cellules dépendantes.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-191">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="8bd8f-192">C’est pourquoi, un trop grand nombre de dépendances de fonctions volatiles risque de ralentir les calculs. Nous vous recommandons d’en utiliser aussi peu que possible.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-192">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="8bd8f-193">Les fonctions de diffusion en continu ne peuvent pas être volatiles.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-193">Streaming functions cannot be volatile.</span></span>

---

## <a name="types"></a><span data-ttu-id="8bd8f-194">Types</span><span class="sxs-lookup"><span data-stu-id="8bd8f-194">Types</span></span>

<span data-ttu-id="8bd8f-195">En spécifiant un type de paramètre, Excel convertit les valeurs en ce type, avant d’appeler la fonction.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-195">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="8bd8f-196">Si le type est `any`, Excel n’effectue pas de conversion.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-196">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="8bd8f-197">Types de valeur</span><span class="sxs-lookup"><span data-stu-id="8bd8f-197">Value types</span></span>

<span data-ttu-id="8bd8f-198">Une valeur unique peut être représentée à l’aide d’un des types suivants : `boolean`, `number` ou `string`.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-198">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="8bd8f-199">Type matrice</span><span class="sxs-lookup"><span data-stu-id="8bd8f-199">Matrix type</span></span>

<span data-ttu-id="8bd8f-200">Utilisez une matrice à deux dimensions pour que le paramètre ou la valeur renvoyée soit une matrice de valeurs.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-200">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="8bd8f-201">Par exemple, le type `number[][]` indique une matrice de nombres.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-201">For example, the type `number[][]` indicates a matrix of numbers.</span></span> <span data-ttu-id="8bd8f-202">`string[][]`indique une matrice de chaînes.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-202">`string[][]` indicates a matrix of strings.</span></span> 

### <a name="error-type"></a><span data-ttu-id="8bd8f-203">Type d’erreur</span><span class="sxs-lookup"><span data-stu-id="8bd8f-203">Error type</span></span>

<span data-ttu-id="8bd8f-204">Une fonction qui n’est pas une fonction de diffusion en continu peut indiquer une erreur en renvoyant un type Error.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-204">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="8bd8f-205">Une fonction de diffusion en continu peut indiquer une erreur en appelant la méthode setResult() avec un type Error.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-205">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="8bd8f-206">Promise</span><span class="sxs-lookup"><span data-stu-id="8bd8f-206">Promise</span></span>

<span data-ttu-id="8bd8f-207">Une fonction peut renvoyer un objet Promise (pour « promesse »). Ce dernier fournit une valeur lors de sa résolution.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-207">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="8bd8f-208">Si la résolution de l’objet Promise est refusée, cela entraîne une erreur.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-208">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="8bd8f-209">Autres types</span><span class="sxs-lookup"><span data-stu-id="8bd8f-209">Other types</span></span>

<span data-ttu-id="8bd8f-210">Tout autre type sera traité comme une erreur.</span><span class="sxs-lookup"><span data-stu-id="8bd8f-210">Any other type will be treated as an error.</span></span>

## <a name="see-also"></a><span data-ttu-id="8bd8f-211">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8bd8f-211">See also</span></span>

* [<span data-ttu-id="8bd8f-212">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="8bd8f-212">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="8bd8f-213">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="8bd8f-213">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="8bd8f-214">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="8bd8f-214">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="8bd8f-215">Fonctions personnalisées changelog</span><span class="sxs-lookup"><span data-stu-id="8bd8f-215">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="8bd8f-216">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="8bd8f-216">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="8bd8f-217">Débogage des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="8bd8f-217">Custom functions debugging</span></span>](custom-functions-debugging.md)