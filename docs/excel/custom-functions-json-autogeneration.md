---
ms.date: 05/03/2019
description: Utiliser les balises JSDOC pour créer dynamiquement vos fonctions personnalisées de métadonnées JSON.
title: Générer automatiquement des métadonnées JSON pour des fonctions personnalisées
localization_priority: Priority
ms.openlocfilehash: 67026e7c19580c3420638b4f37e333e50fce1b44
ms.sourcegitcommit: b299b8a5dfffb6102cb14b431bdde4861abfb47f
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/30/2019
ms.locfileid: "34589131"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a><span data-ttu-id="a8458-103">Générer automatiquement des métadonnées JSON pour des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="a8458-103">Autogenerate JSON metadata for custom functions</span></span>

<span data-ttu-id="a8458-104">Si vous écrivez une fonction Excel personnalisée en JavaScript ou TypeScript, vous pouvez utiliser les balises JSDoc pour la détailler en ajoutant des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="a8458-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="a8458-105">Les balises JSDoc sont ensuite utilisées lors de la génération pour créer le [fichier de métadonnées JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="a8458-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="a8458-106">En utilisant des balises JSDoc, vous n’avez plus besoin de modifier manuellement le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="a8458-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="a8458-107">Ajoutez la balise `@customfunction` dans les commentaires du code d’une fonction JavaScript ou TypeScript pour indiquer qu’il s’agit d’une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="a8458-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="a8458-108">Vous pouvez fournir les types de paramètres de la fonction en utilisant la balise[@param](#param)dans JavaScript, ou en précisant le [type de fonction](https://www.typescriptlang.org/docs/handbook/functions.html) dans TypeScript.</span><span class="sxs-lookup"><span data-stu-id="a8458-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="a8458-109">Si vous voulez en savoir plus, veuillez consulter les sections relatives à la balise[@param](#param) et aux sections[types](#types).</span><span class="sxs-lookup"><span data-stu-id="a8458-109">For more information, see the [@param](#param) tag and [Types](#types) section.</span></span>

## <a name="jsdoc-tags"></a><span data-ttu-id="a8458-110">Balises JSDoc</span><span class="sxs-lookup"><span data-stu-id="a8458-110">JSDoc Tags</span></span>
<span data-ttu-id="a8458-111">Voici quelles sont les balises JSDoc prises en charge dans les fonctions Excel personnalisées :</span><span class="sxs-lookup"><span data-stu-id="a8458-111">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [<span data-ttu-id="a8458-112">@ annulable</span><span class="sxs-lookup"><span data-stu-id="a8458-112">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="a8458-113">[@fonctionpersonnalisée](#customfunction)nom id</span><span class="sxs-lookup"><span data-stu-id="a8458-113">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="a8458-114">url[@urlaide](#helpurl)</span><span class="sxs-lookup"><span data-stu-id="a8458-114">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="a8458-115">[@param](#param) _{type}_ description nom</span><span class="sxs-lookup"><span data-stu-id="a8458-115">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="a8458-116">@requièreuneadresse</span><span class="sxs-lookup"><span data-stu-id="a8458-116">@requiresAddress</span></span>](#requiresAddress)
* <span data-ttu-id="a8458-117">[@renvoie](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="a8458-117">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="a8458-118">@diffusionencontinu</span><span class="sxs-lookup"><span data-stu-id="a8458-118">@streaming</span></span>](#streaming)
* [<span data-ttu-id="a8458-119">@volatile</span><span class="sxs-lookup"><span data-stu-id="a8458-119">@volatile</span></span>](#volatile)

---
### <a name="cancelable"></a><span data-ttu-id="a8458-120">@ annulable</span><span class="sxs-lookup"><span data-stu-id="a8458-120">@cancelable</span></span>
<a id="cancelable"/>

<span data-ttu-id="a8458-121">Indique qu’une fonction personnalisée souhaite effectuer une action lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="a8458-121">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="a8458-122">Le dernier paramètre de la fonction doit être de type `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="a8458-122">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="a8458-123">La fonction peut attribuer une fonction à la propriété `oncanceled` pour désigner l’action à effectuer lors de l’annulation de la fonction.</span><span class="sxs-lookup"><span data-stu-id="a8458-123">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="a8458-124">Si le dernier paramètre de fonction est de type `CustomFunctions.CancelableInvocation`, il sera considéré comme `@cancelable`, même si la balise n’apparaît pas.</span><span class="sxs-lookup"><span data-stu-id="a8458-124">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="a8458-125">Une fonction ne peut pas contenir les deux balises `@cancelable` et `@streaming`.</span><span class="sxs-lookup"><span data-stu-id="a8458-125">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a><span data-ttu-id="a8458-126">@fonctionpersonnalisée</span><span class="sxs-lookup"><span data-stu-id="a8458-126">@customfunction</span></span>
<a id="customfunction"/>

<span data-ttu-id="a8458-127">Syntaxe: @fonctionpersonnalisée_id_ _nom_</span><span class="sxs-lookup"><span data-stu-id="a8458-127">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="a8458-128">Spécifiez cette balise pour traiter la fonction JavaScript/TypeScript comme une fonction Excel personnalisée.</span><span class="sxs-lookup"><span data-stu-id="a8458-128">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="a8458-129">Cette balise est requise pour créer des métadonnées pour la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="a8458-129">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="a8458-130">Vous devez également insérer un appel vers`CustomFunctions.associate("id", functionName);`</span><span class="sxs-lookup"><span data-stu-id="a8458-130">There should also be a call to `CustomFunctions.associate("id", functionName);`</span></span>

#### <a name="id"></a><span data-ttu-id="a8458-131">id</span><span class="sxs-lookup"><span data-stu-id="a8458-131">id</span></span>

<span data-ttu-id="a8458-132">L’id est utilisé en tant qu’identificateur invariant pour la fonction personnalisée stockée dans le document.</span><span class="sxs-lookup"><span data-stu-id="a8458-132">The id is used as the invariant identifier for the custom function stored in the document.</span></span> <span data-ttu-id="a8458-133">Elle ne doit pas changer.</span><span class="sxs-lookup"><span data-stu-id="a8458-133">It should not change.</span></span>

* <span data-ttu-id="a8458-134">Si l’id n’est pas fourni, le nom de la fonction JavaScript/TypeScript est convertie en majuscules, et les caractères rejetés sont supprimés.</span><span class="sxs-lookup"><span data-stu-id="a8458-134">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="a8458-135">L’id doit être unique pour toutes les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="a8458-135">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="a8458-136">Les caractères autorisés sont les suivants : A-Z, a-z, 0-9, traits de soulignement (\_) et point (.).</span><span class="sxs-lookup"><span data-stu-id="a8458-136">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

#### <a name="name"></a><span data-ttu-id="a8458-137">name</span><span class="sxs-lookup"><span data-stu-id="a8458-137">name</span></span>

<span data-ttu-id="a8458-138">Fournit le nom d’affichage de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="a8458-138">Provides the display name for the custom function.</span></span>

* <span data-ttu-id="a8458-139">Si aucun nom n’est fourni, l’id servira aussi de nom.</span><span class="sxs-lookup"><span data-stu-id="a8458-139">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="a8458-140">Caractères autorisés : [caractères alphanumériques Unicode](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic) (lettres, chiffres), point (.) et trait de soulignement (\_).</span><span class="sxs-lookup"><span data-stu-id="a8458-140">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="a8458-141">Doit commencer par une lettre.</span><span class="sxs-lookup"><span data-stu-id="a8458-141">Must start with a letter.</span></span>
* <span data-ttu-id="a8458-142">Sa longueur maximale est limitée à 128 caractères.</span><span class="sxs-lookup"><span data-stu-id="a8458-142">Maximum length is 128 characters.</span></span>

---
### <a name="helpurl"></a><span data-ttu-id="a8458-143">@urlaide</span><span class="sxs-lookup"><span data-stu-id="a8458-143">@helpurl</span></span>
<a id="helpurl"/>

<span data-ttu-id="a8458-144">Syntaxe: @urlaide_url_</span><span class="sxs-lookup"><span data-stu-id="a8458-144">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="a8458-145">L’_url_ fournie est affichée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="a8458-145">The provided _url_ is displayed in Excel.</span></span>

---
### <a name="param"></a><span data-ttu-id="a8458-146">@param</span><span class="sxs-lookup"><span data-stu-id="a8458-146">@param</span></span>
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="a8458-147">JavaScript</span><span class="sxs-lookup"><span data-stu-id="a8458-147">JavaScript</span></span>

<span data-ttu-id="a8458-148">Syntaxe JavaScript : @param {type} nom_description_</span><span class="sxs-lookup"><span data-stu-id="a8458-148">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="a8458-149">`{type}`doit spécifier les informations de type entre deux accolades.</span><span class="sxs-lookup"><span data-stu-id="a8458-149">`{type}` should specify the type info within curly braces.</span></span> <span data-ttu-id="a8458-150">Consultez la section [Types](##types) pour savoir quels types peuvent être utilisés.</span><span class="sxs-lookup"><span data-stu-id="a8458-150">See the [Types](##types) for more information about the types which may be used.</span></span> <span data-ttu-id="a8458-151">Facultatif : si aucun serveur n’est spécifié, le type `any` sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="a8458-151">Optional: if not specified, the type `any` will be used.</span></span>
* <span data-ttu-id="a8458-152">`name`spécifie le paramètre auquel s’applique la balise.</span><span class="sxs-lookup"><span data-stu-id="a8458-152">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="a8458-153">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="a8458-153">Required.</span></span>
* <span data-ttu-id="a8458-154">`description`fournit la description qui s’affiche dans Excel pour le paramètre de la fonction.</span><span class="sxs-lookup"><span data-stu-id="a8458-154">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="a8458-155">Facultatif.</span><span class="sxs-lookup"><span data-stu-id="a8458-155">Optional.</span></span>

<span data-ttu-id="a8458-156">Pour désigner un paramètre de fonction personnalisée comme étant facultatif :</span><span class="sxs-lookup"><span data-stu-id="a8458-156">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="a8458-157">Placez les crochets autour du nom du paramètre.</span><span class="sxs-lookup"><span data-stu-id="a8458-157">Put square brackets around the parameter name.</span></span> <span data-ttu-id="a8458-158">Par exemple : `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="a8458-158">For example: `@param {string} [text] Optional text`.</span></span>

> [!NOTE]
> <span data-ttu-id="a8458-159">La valeur par défaut pour les paramètres facultatifs est `null`.</span><span class="sxs-lookup"><span data-stu-id="a8458-159">The default value for optional parameters is `null`.</span></span>

#### <a name="typescript"></a><span data-ttu-id="a8458-160">TypeScript</span><span class="sxs-lookup"><span data-stu-id="a8458-160">TypeScript</span></span>

<span data-ttu-id="a8458-161">Syntaxe TypeScript : nom @param_description_</span><span class="sxs-lookup"><span data-stu-id="a8458-161">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="a8458-162">`name`spécifie le paramètre auquel s’applique la balise.</span><span class="sxs-lookup"><span data-stu-id="a8458-162">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="a8458-163">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="a8458-163">Required.</span></span>
* <span data-ttu-id="a8458-164">`description`fournit la description qui s’affiche dans Excel pour le paramètre de la fonction.</span><span class="sxs-lookup"><span data-stu-id="a8458-164">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="a8458-165">Facultatif.</span><span class="sxs-lookup"><span data-stu-id="a8458-165">Optional.</span></span>

<span data-ttu-id="a8458-166">Consultez la section [Types](##types) pour savoir quels types de paramètres de fonction peuvent être utilisés.</span><span class="sxs-lookup"><span data-stu-id="a8458-166">See the [Types](##types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="a8458-167">Pour désigner un paramètre de fonction personnalisée comme étant facultatif, effectuez l’une des actions suivantes :</span><span class="sxs-lookup"><span data-stu-id="a8458-167">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="a8458-168">Utilisez un paramètre facultatif.</span><span class="sxs-lookup"><span data-stu-id="a8458-168">Use an optional parameter.</span></span> <span data-ttu-id="a8458-169">Par exemple : `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="a8458-169">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="a8458-170">Définissez ce paramètre sur une valeur par défaut.</span><span class="sxs-lookup"><span data-stu-id="a8458-170">Give the parameter a default value.</span></span> <span data-ttu-id="a8458-171">Par exemple : `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="a8458-171">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="a8458-172">Pour consulter une description détaillée du @param, reportez-vous à la page suivante : [JSDoc](https://usejsdoc.org/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="a8458-172">For detailed description of the @param see: [JSDoc](https://usejsdoc.org/tags-param.html)</span></span>

> [!NOTE]
> <span data-ttu-id="a8458-173">La valeur par défaut pour les paramètres facultatifs est `null`.</span><span class="sxs-lookup"><span data-stu-id="a8458-173">The default value for optional parameters is `null`.</span></span>

---
### <a name="requiresaddress"></a><span data-ttu-id="a8458-174">@requièreuneadresse</span><span class="sxs-lookup"><span data-stu-id="a8458-174">@requiresAddress</span></span>
<a id="requiresAddress"/>

<span data-ttu-id="a8458-175">Indique que l’adresse de la cellule dans laquelle la fonction est évaluée doit être fournie.</span><span class="sxs-lookup"><span data-stu-id="a8458-175">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span> 

<span data-ttu-id="a8458-176">Le dernier paramètre de la fonction doit être de type `CustomFunctions.Invocation` ou un type dérivé.</span><span class="sxs-lookup"><span data-stu-id="a8458-176">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="a8458-177">Lorsque la fonction est appelée, la propriété `address` contiendra l’adresse.</span><span class="sxs-lookup"><span data-stu-id="a8458-177">When the function is called, the `address` property will contain the address.</span></span>

---
### <a name="returns"></a><span data-ttu-id="a8458-178">@renvoie :</span><span class="sxs-lookup"><span data-stu-id="a8458-178">@returns</span></span>
<a id="returns"/>

<span data-ttu-id="a8458-179">Syntaxe: @renvoie {_type_}</span><span class="sxs-lookup"><span data-stu-id="a8458-179">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="a8458-180">Fournit le type pour la valeur renvoyée.</span><span class="sxs-lookup"><span data-stu-id="a8458-180">Provides the type for the return value.</span></span>

<span data-ttu-id="a8458-181">Si `{type}` est omis, les informations de type TypeScript seront utilisées.</span><span class="sxs-lookup"><span data-stu-id="a8458-181">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="a8458-182">S’il n’existe aucune information définissant le type, ce dernier sera `any`.</span><span class="sxs-lookup"><span data-stu-id="a8458-182">If there is no type info, the type will be `any`.</span></span>

---
### <a name="streaming"></a><span data-ttu-id="a8458-183">@diffusionencontinu</span><span class="sxs-lookup"><span data-stu-id="a8458-183">@streaming</span></span>
<a id="streaming"/>

<span data-ttu-id="a8458-184">Utilisé pour indiquer qu’une fonction personnalisée est une fonction diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="a8458-184">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="a8458-185">Le dernier paramètre doit être de type `CustomFunctions.StreamingInvocation<ResultType>`.</span><span class="sxs-lookup"><span data-stu-id="a8458-185">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="a8458-186">La fonction doit renvoyer `void`.</span><span class="sxs-lookup"><span data-stu-id="a8458-186">The function should return `void`.</span></span>

<span data-ttu-id="a8458-187">Les fonctions de diffusion en continu ne renvoient pas de valeurs directement, mais doivent plutôt appeler `setResult(result: ResultType)` en utilisant le dernier paramètre.</span><span class="sxs-lookup"><span data-stu-id="a8458-187">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="a8458-188">Les exceptions levées par une fonction en continu sont ignorées.</span><span class="sxs-lookup"><span data-stu-id="a8458-188">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="a8458-189">`setResult()`peut être appelée avec Error pour indiquer un résultat erroné.</span><span class="sxs-lookup"><span data-stu-id="a8458-189">`setResult()` may be called with Error to indicate an error result.</span></span>

<span data-ttu-id="a8458-190">Vous ne pouvez pas utiliser les balises en diffusion en continu comme [@volatile](#volatile).</span><span class="sxs-lookup"><span data-stu-id="a8458-190">Streaming functions cannot be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a><span data-ttu-id="a8458-191">@volatile</span><span class="sxs-lookup"><span data-stu-id="a8458-191">@volatile</span></span>
<a id="volatile"/>

<span data-ttu-id="a8458-192">Une fonction volatile est une fonction dont le résultat peut changer d’un moment à l’autre, même si elle ne récupère pas d’argument ou si ses arguments ne changent pas.</span><span class="sxs-lookup"><span data-stu-id="a8458-192">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="a8458-193">À chaque calcul, Excel réévalue les cellules contenant des fonctions volatiles, ainsi que toutes leurs cellules dépendantes.</span><span class="sxs-lookup"><span data-stu-id="a8458-193">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="a8458-194">C’est pourquoi, un trop grand nombre de dépendances de fonctions volatiles risque de ralentir les calculs. Nous vous recommandons d’en utiliser aussi peu que possible.</span><span class="sxs-lookup"><span data-stu-id="a8458-194">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="a8458-195">Les fonctions de diffusion en continu ne peuvent pas être volatiles.</span><span class="sxs-lookup"><span data-stu-id="a8458-195">Streaming functions cannot be volatile.</span></span>

---

## <a name="types"></a><span data-ttu-id="a8458-196">Types</span><span class="sxs-lookup"><span data-stu-id="a8458-196">Types</span></span>

<span data-ttu-id="a8458-197">En spécifiant un type de paramètre, Excel convertit les valeurs en ce type, avant d’appeler la fonction.</span><span class="sxs-lookup"><span data-stu-id="a8458-197">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="a8458-198">Si le type est `any`, Excel n’effectue pas de conversion.</span><span class="sxs-lookup"><span data-stu-id="a8458-198">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="a8458-199">Types de valeur</span><span class="sxs-lookup"><span data-stu-id="a8458-199">Value types</span></span>

<span data-ttu-id="a8458-200">Une valeur unique peut être représentée à l’aide d’un des types suivants : `boolean`, `number` ou `string`.</span><span class="sxs-lookup"><span data-stu-id="a8458-200">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="a8458-201">Type matrice</span><span class="sxs-lookup"><span data-stu-id="a8458-201">Matrix type</span></span>

<span data-ttu-id="a8458-202">Utilisez une matrice à deux dimensions pour que le paramètre ou la valeur renvoyée soit une matrice de valeurs.</span><span class="sxs-lookup"><span data-stu-id="a8458-202">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="a8458-203">Par exemple, le type `number[][]` indique une matrice de nombres.</span><span class="sxs-lookup"><span data-stu-id="a8458-203">For example, the type `number[][]` indicates a matrix of numbers.</span></span> <span data-ttu-id="a8458-204">`string[][]`indique une matrice de chaînes.</span><span class="sxs-lookup"><span data-stu-id="a8458-204">`string[][]` indicates a matrix of strings.</span></span> 

### <a name="error-type"></a><span data-ttu-id="a8458-205">Type d’erreur</span><span class="sxs-lookup"><span data-stu-id="a8458-205">Error type</span></span>

<span data-ttu-id="a8458-206">Une fonction qui n’est pas une fonction de diffusion en continu peut indiquer une erreur en renvoyant un type Error.</span><span class="sxs-lookup"><span data-stu-id="a8458-206">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="a8458-207">Une fonction de diffusion en continu peut indiquer une erreur en appelant la méthode setResult() avec un type Error.</span><span class="sxs-lookup"><span data-stu-id="a8458-207">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="a8458-208">Promise</span><span class="sxs-lookup"><span data-stu-id="a8458-208">Promise</span></span>

<span data-ttu-id="a8458-209">Une fonction peut renvoyer un objet Promise (pour « promesse »). Ce dernier fournit une valeur lors de sa résolution.</span><span class="sxs-lookup"><span data-stu-id="a8458-209">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="a8458-210">Si la résolution de l’objet Promise est refusée, cela entraîne une erreur.</span><span class="sxs-lookup"><span data-stu-id="a8458-210">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="a8458-211">Autres types</span><span class="sxs-lookup"><span data-stu-id="a8458-211">Other types</span></span>

<span data-ttu-id="a8458-212">Tout autre type sera traité comme une erreur.</span><span class="sxs-lookup"><span data-stu-id="a8458-212">Any other type will be treated as an error.</span></span>

## <a name="next-steps"></a><span data-ttu-id="a8458-213">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="a8458-213">Next steps</span></span>
<span data-ttu-id="a8458-214">Découvrez les [conventions d’affectation des noms des fonctions personnalisées](custom-functions-naming.md).</span><span class="sxs-lookup"><span data-stu-id="a8458-214">Learn about [naming conventions for custom functions](custom-functions-naming.md).</span></span> <span data-ttu-id="a8458-215">Découvrez également comment [localiser vos fonctions](custom-functions-localize.md), ce qui implique que vous [écriviez votre fichier JSON à la main](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="a8458-215">Alternatively, learn how to [localize your functions](custom-functions-localize.md) which requires you to [write your JSON file by hand](custom-functions-json.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="a8458-216">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a8458-216">See also</span></span>

* [<span data-ttu-id="a8458-217">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="a8458-217">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="a8458-218">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="a8458-218">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="a8458-219">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="a8458-219">Create custom functions in Excel</span></span>](custom-functions-overview.md)
