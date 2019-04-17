---
ms.date: 04/03/2019
description: Utiliser les balises JSDOC pour créer dynamiquement vos fonctions personnalisées de métadonnées JSON.
title: Créer les métadonnées JSON pour des fonctions personnalisées (aperçu)
localization_priority: Priority
ms.openlocfilehash: c6d89684da2d0773ccfb1763e5e3e426e647523b
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/04/2019
ms.locfileid: "31478955"
---
# <a name="create-json-metadata-for-custom-functions-preview"></a><span data-ttu-id="dd78e-103">Créer les métadonnées JSON pour des fonctions personnalisées (aperçu)</span><span class="sxs-lookup"><span data-stu-id="dd78e-103">Create JSON metadata for custom functions (preview)</span></span>

<span data-ttu-id="dd78e-104">Si vous écrivez une fonction Excel personnalisée en JavaScript ou TypeScript, vous pouvez utiliser les balises JSDoc pour la détailler en ajoutant des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="dd78e-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="dd78e-105">Les balises JSDoc sont ensuite utilisées lors de la génération pour créer le [fichier de métadonnées JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="dd78e-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="dd78e-106">En utilisant des balises JSDoc, vous n’avez plus besoin de modifier manuellement le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="dd78e-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

<span data-ttu-id="dd78e-107">Ajoutez la balise `@customfunction` dans les commentaires du code d’une fonction JavaScript ou TypeScript pour indiquer qu’il s’agit d’une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="dd78e-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="dd78e-108">Vous pouvez fournir les types de paramètres de la fonction en utilisant la balise [@param](#param) dans JavaScript, ou en précisant le [type de fonction](http://www.typescriptlang.org/docs/handbook/functions.html) dans TypeScript.</span><span class="sxs-lookup"><span data-stu-id="dd78e-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](http://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="dd78e-109">Si vous voulez en savoir plus, veuillez consulter les sections relatives à la balise [@param](#param) et aux [types](#Types).</span><span class="sxs-lookup"><span data-stu-id="dd78e-109">For more information, see the [@param](#param) tag and [Types](#Types) section.</span></span>

## <a name="jsdoc-tags"></a><span data-ttu-id="dd78e-110">Balises JSDoc</span><span class="sxs-lookup"><span data-stu-id="dd78e-110">JSDoc Tags</span></span>
<span data-ttu-id="dd78e-111">Voici quelles sont les balises JSDoc prises en charge dans les fonctions Excel personnalisées :</span><span class="sxs-lookup"><span data-stu-id="dd78e-111">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [@cancelable](#cancelable)
* <span data-ttu-id="dd78e-112">[@customfunction](#customfunction) id name</span><span class="sxs-lookup"><span data-stu-id="dd78e-112">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="dd78e-113">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="dd78e-113">URL</span></span>
* <span data-ttu-id="dd78e-114">[@param](#param) _{type}_ name description</span><span class="sxs-lookup"><span data-stu-id="dd78e-114">[@param](#param) _{type}_ name description</span></span>
* [@requiresAddress](#requiresAddress)
* <span data-ttu-id="dd78e-115">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="dd78e-115">Type</span></span>
* [@streaming](#streaming)
* [@volatile](#volatile)

---
### <a name="cancelable"></a>@cancelable
<a id="cancelable"/>

<span data-ttu-id="dd78e-116">Indique qu’une fonction personnalisée souhaite effectuer une action lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="dd78e-116">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="dd78e-117">Le dernier paramètre de la fonction doit être de type `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="dd78e-117">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="dd78e-118">La fonction peut attribuer une fonction à la propriété `oncanceled` pour désigner l’action à effectuer lors de l’annulation de la fonction.</span><span class="sxs-lookup"><span data-stu-id="dd78e-118">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="dd78e-119">Si le dernier paramètre de fonction est de type `CustomFunctions.CancelableInvocation`, il sera considéré comme `@cancelable`, même si la balise n’apparaît pas.</span><span class="sxs-lookup"><span data-stu-id="dd78e-119">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="dd78e-120">Une fonction ne peut pas contenir les deux balises `@cancelable` et `@streaming`.</span><span class="sxs-lookup"><span data-stu-id="dd78e-120">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a>@customfunction
<a id="customfunction"/>

<span data-ttu-id="dd78e-121">Syntaxe : @customfunction _id_ _name_</span><span class="sxs-lookup"><span data-stu-id="dd78e-121">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="dd78e-122">Spécifiez cette balise pour traiter la fonction JavaScript/TypeScript comme une fonction Excel personnalisée.</span><span class="sxs-lookup"><span data-stu-id="dd78e-122">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="dd78e-123">Cette balise est requise pour créer des métadonnées pour la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="dd78e-123">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="dd78e-124">Vous devez également insérer un appel vers</span><span class="sxs-lookup"><span data-stu-id="dd78e-124">There should also be a call to</span></span> `CustomFunctions.associate("id", functionName);`

#### <a name="id"></a><span data-ttu-id="dd78e-125">id</span><span class="sxs-lookup"><span data-stu-id="dd78e-125">id</span></span> 

<span data-ttu-id="dd78e-126">L’id est utilisé en tant qu’identificateur invariant pour la fonction personnalisée stockée dans le document.</span><span class="sxs-lookup"><span data-stu-id="dd78e-126">The id is used as the invariant identifier for the custom function stored in the document.</span></span> <span data-ttu-id="dd78e-127">Elle ne doit pas changer.</span><span class="sxs-lookup"><span data-stu-id="dd78e-127">It should not change.</span></span>

* <span data-ttu-id="dd78e-128">Si l’id n’est pas fourni, le nom de la fonction JavaScript/TypeScript est convertie en majuscules, et les caractères rejetés sont supprimés.</span><span class="sxs-lookup"><span data-stu-id="dd78e-128">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="dd78e-129">L’id doit être unique pour toutes les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="dd78e-129">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="dd78e-130">Seuls les caractères alphanumériques majuscules et minuscules (A-Z, a-z, 0-9) et le point (.) sont autorisés.</span><span class="sxs-lookup"><span data-stu-id="dd78e-130">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

#### <a name="name"></a><span data-ttu-id="dd78e-131">name</span><span class="sxs-lookup"><span data-stu-id="dd78e-131">name</span></span>

<span data-ttu-id="dd78e-132">Fournit le nom d’affichage de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="dd78e-132">Provides the display name of a custom category for the property.</span></span> 

* <span data-ttu-id="dd78e-133">Si aucun nom n’est fourni, l’id servira aussi de nom.</span><span class="sxs-lookup"><span data-stu-id="dd78e-133">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="dd78e-134">Caractères autorisés : [caractères alphanumériques Unicode](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic) (lettres, chiffres), point (.) et trait de soulignement (\_).</span><span class="sxs-lookup"><span data-stu-id="dd78e-134">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="dd78e-135">Doit commencer par une lettre.</span><span class="sxs-lookup"><span data-stu-id="dd78e-135">Must start with a letter.</span></span>
* <span data-ttu-id="dd78e-136">Sa longueur maximale est limitée à 128 caractères.</span><span class="sxs-lookup"><span data-stu-id="dd78e-136">Maximum length is 255 characters.</span></span>

---
### <a name="helpurl"></a>@helpurl
<a id="helpurl"/>

<span data-ttu-id="dd78e-137">Syntaxe : @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="dd78e-137">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="dd78e-138">L’_url_ fournie est affichée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="dd78e-138">The provided _url_ is displayed in Excel.</span></span>

---
### <a name="param"></a>@param
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="dd78e-139">JavaScript</span><span class="sxs-lookup"><span data-stu-id="dd78e-139">JavaScript</span></span>

<span data-ttu-id="dd78e-140">Syntaxe JavaScript : @param {type} name _description_</span><span class="sxs-lookup"><span data-stu-id="dd78e-140">JavaScript Syntax: @param {type} name _description_</span></span>

* `{type}` <span data-ttu-id="dd78e-141">doit spécifier les informations de type entre deux accolades.</span><span class="sxs-lookup"><span data-stu-id="dd78e-141">should specify the type info within curly braces.</span></span> <span data-ttu-id="dd78e-142">Consultez la section [Types](##types) pour savoir quels types peuvent être utilisés.</span><span class="sxs-lookup"><span data-stu-id="dd78e-142">See the [Types](##types) for more information about the types which may be used.</span></span> <span data-ttu-id="dd78e-143">Facultatif : si aucun serveur n’est spécifié, le type `any` sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="dd78e-143">Optional: if not specified, the type `any` will be used.</span></span>
* `name` <span data-ttu-id="dd78e-144">spécifie le paramètre auquel s’applique la balise @param.</span><span class="sxs-lookup"><span data-stu-id="dd78e-144">specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="dd78e-145">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="dd78e-145">Required.</span></span>
* `description` <span data-ttu-id="dd78e-146">fournit la description qui s’affiche dans Excel pour le paramètre de la fonction.</span><span class="sxs-lookup"><span data-stu-id="dd78e-146">provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="dd78e-147">Facultatif.</span><span class="sxs-lookup"><span data-stu-id="dd78e-147">Optional.</span></span>

<span data-ttu-id="dd78e-148">Pour désigner un paramètre de fonction personnalisée comme étant facultatif :</span><span class="sxs-lookup"><span data-stu-id="dd78e-148">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="dd78e-149">Placez les crochets autour du nom du paramètre.</span><span class="sxs-lookup"><span data-stu-id="dd78e-149">Put square brackets around the parameter name.</span></span> <span data-ttu-id="dd78e-150">Par exemple : `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="dd78e-150">For example: `@param {string} [text] Optional text`.</span></span>

#### <a name="typescript"></a><span data-ttu-id="dd78e-151">TypeScript</span><span class="sxs-lookup"><span data-stu-id="dd78e-151">TypeScript</span></span>

<span data-ttu-id="dd78e-152">Syntaxe TypeScript : @param name _description_</span><span class="sxs-lookup"><span data-stu-id="dd78e-152">TypeScript Syntax: @param name _description_</span></span>

* `name` <span data-ttu-id="dd78e-153">spécifie le paramètre auquel s’applique la balise @param.</span><span class="sxs-lookup"><span data-stu-id="dd78e-153">specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="dd78e-154">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="dd78e-154">Required.</span></span>
* `description` <span data-ttu-id="dd78e-155">fournit la description qui s’affiche dans Excel pour le paramètre de la fonction.</span><span class="sxs-lookup"><span data-stu-id="dd78e-155">provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="dd78e-156">Facultatif.</span><span class="sxs-lookup"><span data-stu-id="dd78e-156">Optional.</span></span>

<span data-ttu-id="dd78e-157">Consultez la section [Types](##types) pour savoir quels types de paramètres de fonction peuvent être utilisés.</span><span class="sxs-lookup"><span data-stu-id="dd78e-157">See the [Types](##types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="dd78e-158">Pour désigner un paramètre de fonction personnalisée comme étant facultatif, effectuez l’une des actions suivantes :</span><span class="sxs-lookup"><span data-stu-id="dd78e-158">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="dd78e-159">Utilisez un paramètre facultatif.</span><span class="sxs-lookup"><span data-stu-id="dd78e-159">Use an optional parameter.</span></span> <span data-ttu-id="dd78e-160">Par exemple :</span><span class="sxs-lookup"><span data-stu-id="dd78e-160">For example:</span></span> `function f(text?: string)`
* <span data-ttu-id="dd78e-161">Définissez ce paramètre sur une valeur par défaut.</span><span class="sxs-lookup"><span data-stu-id="dd78e-161">Give the parameter a default value.</span></span> <span data-ttu-id="dd78e-162">Par exemple :</span><span class="sxs-lookup"><span data-stu-id="dd78e-162">For example:</span></span> `function f(text: string = "abc")`

<span data-ttu-id="dd78e-163">Pour consulter une description détaillée du paramètre @param, reportez-vous à la page suivante : [JSDoc](http://usejsdoc.org/tags-param.html) (contenu en anglais)</span><span class="sxs-lookup"><span data-stu-id="dd78e-163">For a detailed description of the code, see "HelloData Details."</span></span>

---
### <a name="requiresaddress"></a>@requiresAddress
<a id="requiresAddress"/>

<span data-ttu-id="dd78e-164">Indique que l’adresse de la cellule dans laquelle la fonction est évaluée doit être fournie.</span><span class="sxs-lookup"><span data-stu-id="dd78e-164">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span> 

<span data-ttu-id="dd78e-165">Le dernier paramètre de la fonction doit être de type `CustomFunctions.Invocation` ou un type dérivé.</span><span class="sxs-lookup"><span data-stu-id="dd78e-165">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="dd78e-166">Lorsque la fonction est appelée, la propriété `address` contiendra l’adresse.</span><span class="sxs-lookup"><span data-stu-id="dd78e-166">When the function is called, the `address` property will contain the address.</span></span>

---
### <a name="returns"></a>@returns
<a id="returns"/>

<span data-ttu-id="dd78e-167">Syntaxe : @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="dd78e-167">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="dd78e-168">Fournit le type pour la valeur renvoyée.</span><span class="sxs-lookup"><span data-stu-id="dd78e-168">Provides the type for the return value.</span></span>

<span data-ttu-id="dd78e-169">Si `{type}` est omis, les informations de type TypeScript seront utilisées.</span><span class="sxs-lookup"><span data-stu-id="dd78e-169">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="dd78e-170">S’il n’existe aucune information définissant le type, ce dernier sera `any`.</span><span class="sxs-lookup"><span data-stu-id="dd78e-170">If there is no type info, the type will be `any`.</span></span>

---
### <a name="streaming"></a>@streaming
<a id="streaming"/>

<span data-ttu-id="dd78e-171">Utilisé pour indiquer qu’une fonction personnalisée est une fonction diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="dd78e-171">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="dd78e-172">Le dernier paramètre doit être de type `CustomFunctions.StreamingInvocation<ResultType>`.</span><span class="sxs-lookup"><span data-stu-id="dd78e-172">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="dd78e-173">La fonction doit renvoyer `void`.</span><span class="sxs-lookup"><span data-stu-id="dd78e-173">The function should return `void`.</span></span>

<span data-ttu-id="dd78e-174">Les fonctions de diffusion en continu ne renvoient pas de valeurs directement, mais doivent plutôt appeler `setResult(result: ResultType)` en utilisant le dernier paramètre.</span><span class="sxs-lookup"><span data-stu-id="dd78e-174">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="dd78e-175">Les exceptions levées par une fonction en continu sont ignorées.</span><span class="sxs-lookup"><span data-stu-id="dd78e-175">Exceptions thrown by a streaming function are ignored.</span></span> `setResult()` <span data-ttu-id="dd78e-176">peut être appelée avec Error pour indiquer un résultat erroné.</span><span class="sxs-lookup"><span data-stu-id="dd78e-176">may be called with Error to indicate an error result.</span></span>

<span data-ttu-id="dd78e-177">Vous ne pouvez pas utiliser la balise [@volatile](#volatile) pour les fonctions de diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="dd78e-177">Streaming functions cannot be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a>@volatile
<a id="volatile"/>

<span data-ttu-id="dd78e-178">Une fonction volatile est une fonction dont le résultat peut changer d’un moment à l’autre, même si elle ne récupère pas d’argument ou si ses arguments ne changent pas.</span><span class="sxs-lookup"><span data-stu-id="dd78e-178">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="dd78e-179">À chaque calcul, Excel réévalue les cellules contenant des fonctions volatiles, ainsi que toutes leurs cellules dépendantes.</span><span class="sxs-lookup"><span data-stu-id="dd78e-179">Excel reevaluates cells that contain volatile functions, together with all dependents, every time that it recalculates.</span></span> <span data-ttu-id="dd78e-180">C’est pourquoi, un trop grand nombre de dépendances de fonctions volatiles risque de ralentir les calculs. Nous vous recommandons d’en utiliser aussi peu que possible.</span><span class="sxs-lookup"><span data-stu-id="dd78e-180">For this reason, too much reliance on volatile functions can make recalculation times slow.</span></span>

<span data-ttu-id="dd78e-181">Les fonctions de diffusion en continu ne peuvent pas être volatiles.</span><span class="sxs-lookup"><span data-stu-id="dd78e-181">Streaming functions cannot be volatile.</span></span>

---

## <a name="types"></a><span data-ttu-id="dd78e-182">Types</span><span class="sxs-lookup"><span data-stu-id="dd78e-182">Types</span></span>

<span data-ttu-id="dd78e-183">En spécifiant un type de paramètre, Excel convertit les valeurs en ce type, avant d’appeler la fonction.</span><span class="sxs-lookup"><span data-stu-id="dd78e-183">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="dd78e-184">Si le type est `any`, Excel n’effectue pas de conversion.</span><span class="sxs-lookup"><span data-stu-id="dd78e-184">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="dd78e-185">Types de valeur</span><span class="sxs-lookup"><span data-stu-id="dd78e-185">Value types</span></span>

<span data-ttu-id="dd78e-186">Une valeur unique peut être représentée à l’aide d’un des types suivants : `boolean`, `number` ou `string`.</span><span class="sxs-lookup"><span data-stu-id="dd78e-186">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="dd78e-187">Type matrice</span><span class="sxs-lookup"><span data-stu-id="dd78e-187">Matrix type</span></span>

<span data-ttu-id="dd78e-188">Utilisez une matrice à deux dimensions pour que le paramètre ou la valeur renvoyée soit une matrice de valeurs.</span><span class="sxs-lookup"><span data-stu-id="dd78e-188">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="dd78e-189">Par exemple, le type `number[][]` indique une matrice de nombres.</span><span class="sxs-lookup"><span data-stu-id="dd78e-189">For example, the type `number[][]` indicates a matrix of numbers.</span></span> `string[][]` <span data-ttu-id="dd78e-190">Indique une matrice de chaînes.</span><span class="sxs-lookup"><span data-stu-id="dd78e-190">indicates a matrix of strings.</span></span> 

### <a name="error-type"></a><span data-ttu-id="dd78e-191">Type d’erreur</span><span class="sxs-lookup"><span data-stu-id="dd78e-191">Error Type</span></span>

<span data-ttu-id="dd78e-192">Une fonction qui n’est pas une fonction de diffusion en continu peut indiquer une erreur en renvoyant un type Error.</span><span class="sxs-lookup"><span data-stu-id="dd78e-192">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="dd78e-193">Une fonction de diffusion en continu peut indiquer une erreur en appelant la méthode setResult() avec un type Error.</span><span class="sxs-lookup"><span data-stu-id="dd78e-193">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="dd78e-194">Promise</span><span class="sxs-lookup"><span data-stu-id="dd78e-194">Promise object.</span></span>

<span data-ttu-id="dd78e-195">Une fonction peut renvoyer un objet Promise (pour « promesse »). Ce dernier fournit une valeur lors de sa résolution.</span><span class="sxs-lookup"><span data-stu-id="dd78e-195">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="dd78e-196">Si la résolution de l’objet Promise est refusée, cela entraîne une erreur.</span><span class="sxs-lookup"><span data-stu-id="dd78e-196">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="dd78e-197">Autres types</span><span class="sxs-lookup"><span data-stu-id="dd78e-197">Other solution types</span></span>

<span data-ttu-id="dd78e-198">Tout autre type sera traité comme une erreur.</span><span class="sxs-lookup"><span data-stu-id="dd78e-198">Any other type will be treated as an error.</span></span>

## <a name="see-also"></a><span data-ttu-id="dd78e-199">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="dd78e-199">See also</span></span>

* [<span data-ttu-id="dd78e-200">Métadonnées des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="dd78e-200">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="dd78e-201">Runtime pour les fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="dd78e-201">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="dd78e-202">Meilleures pratiques pour l’utilisation des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="dd78e-202">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="dd78e-203">Journal des modifications des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="dd78e-203">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="dd78e-204">Didacticiel sur les fonctions Excel personnalisées</span><span class="sxs-lookup"><span data-stu-id="dd78e-204">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="dd78e-205">Débogage des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="dd78e-205">Custom functions debugging</span></span>](custom-functions-debugging.md)
