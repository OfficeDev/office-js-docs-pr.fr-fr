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
# <a name="create-json-metadata-for-custom-functions-preview"></a><span data-ttu-id="b620f-103">Créer les métadonnées JSON pour des fonctions personnalisées (aperçu)</span><span class="sxs-lookup"><span data-stu-id="b620f-103">Create JSON metadata for custom functions (preview)</span></span>

<span data-ttu-id="b620f-104">Lorsqu’ une fonction personnalisée Excel est écrite dans JavaScript ou TypeScript, les balises JSDoc servent à fournir des informations supplémentaires sur la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="b620f-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="b620f-105">Les balises JSDoc sont ensuite utilisées au moment de build pour créer le [fichier de métadonnées JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="b620f-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="b620f-106">Utiliser des balises JSDoc vous évite des efforts pour modifier manuellement le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="b620f-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

<span data-ttu-id="b620f-107">Ajouter la`@customfunction` balise dans les commentaires du code d’une fonction JavaScript ou TypeScript pour la marquer comme une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="b620f-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="b620f-108">La fonction types de paramètre peut être fournie à l’aide de la [@param ](#param) balise dans JavaScript, ou via la [fonction type](http://www.typescriptlang.org/docs/handbook/functions.html) dans TypeScript.</span><span class="sxs-lookup"><span data-stu-id="b620f-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](http://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="b620f-109">Pour plus d’informations, consultez la [@param](#param) balise et la section[Types](#Types).</span><span class="sxs-lookup"><span data-stu-id="b620f-109">For more information, see the [@param](#param) tag and [Types](#Types) section.</span></span>

## <a name="jsdoc-tags"></a><span data-ttu-id="b620f-110">Balises JSDoc</span><span class="sxs-lookup"><span data-stu-id="b620f-110">JSDoc Tags</span></span>
<span data-ttu-id="b620f-111">Les balises JSDoc suivants sont prises en charge dans les fonctions personnalisées Excel :</span><span class="sxs-lookup"><span data-stu-id="b620f-111">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [@cancelable](#cancelable)
* <span data-ttu-id="b620f-112">[@customfunction](#customfunction) nom d’ID</span><span class="sxs-lookup"><span data-stu-id="b620f-112">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="b620f-113">[@helpurl](#helpurl)URL</span><span class="sxs-lookup"><span data-stu-id="b620f-113">URL</span></span>
* <span data-ttu-id="b620f-114">[@param](#param) _{type}_ description nom</span><span class="sxs-lookup"><span data-stu-id="b620f-114">[@param](#param) _{type}_ name description</span></span>
* [@requiresAddress](#requiresAddress)
* <span data-ttu-id="b620f-115">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="b620f-115">Type</span></span>
* [@streaming](#streaming)
* [@volatile](#volatile)

---
### <a name="cancelable"></a>@cancelable
<a id="cancelable"/>

<span data-ttu-id="b620f-116">Indique qu’une fonction personnalisée souhaite effectuer une action lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="b620f-116">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="b620f-117">Le dernier paramètre de la fonction doit être de type `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="b620f-117">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="b620f-118">La fonction peut affecter une fonction à la `oncanceled` propriété pour désigner l’action à effectuer lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="b620f-118">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="b620f-119">Si le dernier paramètre de fonction est de type `CustomFunctions.CancelableInvocation`, sera considéré comme `@cancelable` même si la balise n’apparaît pas.</span><span class="sxs-lookup"><span data-stu-id="b620f-119">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="b620f-120">Une fonction ne peut pas contenir les deux balises`@cancelable` et `@streaming`.</span><span class="sxs-lookup"><span data-stu-id="b620f-120">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a>@customfunction
<a id="customfunction"/>

<span data-ttu-id="b620f-121">Syntaxe : @customfunction _id_ _nom_</span><span class="sxs-lookup"><span data-stu-id="b620f-121">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="b620f-122">Spécifier cette balise pour traiter la fonction JavaScript/TypeScript comme une fonction personnalisée Excel.</span><span class="sxs-lookup"><span data-stu-id="b620f-122">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="b620f-123">Cette balise est requise pour créer des métadonnées pour la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="b620f-123">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="b620f-124">Il doit également être un appel au</span><span class="sxs-lookup"><span data-stu-id="b620f-124">There should also be a call to</span></span> `CustomFunctions.associate("id", functionName);`

#### <a name="id"></a><span data-ttu-id="b620f-125">id</span><span class="sxs-lookup"><span data-stu-id="b620f-125">id</span></span> 

<span data-ttu-id="b620f-126">L’id est utilisé en tant qu’identificateur indifférent pour la fonction personnalisée stockée dans le document.</span><span class="sxs-lookup"><span data-stu-id="b620f-126">The id is used as the invariant identifier for the custom function stored in the document.</span></span> <span data-ttu-id="b620f-127">Elle ne doit pas changer.</span><span class="sxs-lookup"><span data-stu-id="b620f-127">It should not change.</span></span>

* <span data-ttu-id="b620f-128">Si l’id n’est pas inclus, le nom de la fonction JavaScript/TypeScript est convertie en majuscules, les caractères rejetés sont supprimés.</span><span class="sxs-lookup"><span data-stu-id="b620f-128">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="b620f-129">L’id doit être unique pour toutes les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="b620f-129">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="b620f-130">Les caractères autorisés sont limités aux : A-Z, a-z, 0-9 et point (.).</span><span class="sxs-lookup"><span data-stu-id="b620f-130">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

#### <a name="name"></a><span data-ttu-id="b620f-131">nom</span><span class="sxs-lookup"><span data-stu-id="b620f-131">name</span></span>

<span data-ttu-id="b620f-132">Fournit le nom d’affichage pour la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="b620f-132">Provides the display name of a custom category for the property.</span></span> 

* <span data-ttu-id="b620f-133">Si aucun nom n’est fourni, l’id est également utilisé comme le nom.</span><span class="sxs-lookup"><span data-stu-id="b620f-133">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="b620f-134">Caractères autorisés : lettres [caractère Unicode alphabétique](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), nombres, point (.) et un trait de soulignement (\_).</span><span class="sxs-lookup"><span data-stu-id="b620f-134">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="b620f-135">Doit commencer par une lettre.</span><span class="sxs-lookup"><span data-stu-id="b620f-135">Must start with a letter.</span></span>
* <span data-ttu-id="b620f-136">La longueur maximale est de 128 caractères.</span><span class="sxs-lookup"><span data-stu-id="b620f-136">Maximum length is 255 characters.</span></span>

---
### <a name="helpurl"></a>@helpurl
<a id="helpurl"/>

<span data-ttu-id="b620f-137">Syntaxe: @helpurl_url_</span><span class="sxs-lookup"><span data-stu-id="b620f-137">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="b620f-138">L’_url_ fournie est affichée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="b620f-138">The provided _url_ is displayed in Excel.</span></span>

---
### <a name="param"></a>@param
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="b620f-139">JavaScript</span><span class="sxs-lookup"><span data-stu-id="b620f-139">JavaScript</span></span>

<span data-ttu-id="b620f-140">Syntaxe JavaScript : @param nom {type} _description_</span><span class="sxs-lookup"><span data-stu-id="b620f-140">JavaScript Syntax: @param {type} name _description_</span></span>

* `{type}` <span data-ttu-id="b620f-141">doit spécifier les informations de type au sein des accolades.</span><span class="sxs-lookup"><span data-stu-id="b620f-141">should specify the type info within curly braces.</span></span> <span data-ttu-id="b620f-142">Voir les[Types](##types) pour plus d’informations sur les types qui peuvent être utilisés.</span><span class="sxs-lookup"><span data-stu-id="b620f-142">See the [Types](##types) for more information about the types which may be used.</span></span> <span data-ttu-id="b620f-143">Facultatif: Si aucun serveur n'est spécifié, le type`any` sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="b620f-143">Optional: if not specified, the type `any` will be used.</span></span>
* `name` <span data-ttu-id="b620f-144">spécifie le paramètre auquel la@parambalise s’applique.</span><span class="sxs-lookup"><span data-stu-id="b620f-144">specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="b620f-145">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="b620f-145">Required.</span></span>
* `description` <span data-ttu-id="b620f-146">fournit la description qui s’affiche dans Excel pour le paramètre de fonction.</span><span class="sxs-lookup"><span data-stu-id="b620f-146">provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="b620f-147">Facultatif.</span><span class="sxs-lookup"><span data-stu-id="b620f-147">Optional.</span></span>

<span data-ttu-id="b620f-148">Pour désigner un paramètre de fonction personnalisée comme étant facultatif:</span><span class="sxs-lookup"><span data-stu-id="b620f-148">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="b620f-149">Placez les crochets autour du paramètre de nom.</span><span class="sxs-lookup"><span data-stu-id="b620f-149">Put square brackets around the parameter name.</span></span> <span data-ttu-id="b620f-150">Par exemple : `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="b620f-150">For example: `@param {string} [text] Optional text`.</span></span>

#### <a name="typescript"></a><span data-ttu-id="b620f-151">TypeScript</span><span class="sxs-lookup"><span data-stu-id="b620f-151">TypeScript</span></span>

<span data-ttu-id="b620f-152">Syntaxe JavaScript : @paramnom_description_</span><span class="sxs-lookup"><span data-stu-id="b620f-152">TypeScript Syntax: @param name _description_</span></span>

* `name` <span data-ttu-id="b620f-153">spécifie le paramètre auquel la@parambalise s’applique.</span><span class="sxs-lookup"><span data-stu-id="b620f-153">specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="b620f-154">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="b620f-154">Required.</span></span>
* `description` <span data-ttu-id="b620f-155">fournit la description qui s’affiche dans Excel pour le paramètre de fonction.</span><span class="sxs-lookup"><span data-stu-id="b620f-155">provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="b620f-156">Facultatif.</span><span class="sxs-lookup"><span data-stu-id="b620f-156">Optional.</span></span>

<span data-ttu-id="b620f-157">Voir les[Types](##types) pour plus d’informations sur les types de paramètre de fonction qui peuvent être utilisés.</span><span class="sxs-lookup"><span data-stu-id="b620f-157">See the [Types](##types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="b620f-158">Pour désigner un paramètre de fonction personnalisée comme étant facultatif, effectuez l’une des actions suivantes:</span><span class="sxs-lookup"><span data-stu-id="b620f-158">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="b620f-159">Utilisez un paramètre facultatif.</span><span class="sxs-lookup"><span data-stu-id="b620f-159">Use an optional parameter.</span></span> <span data-ttu-id="b620f-160">Par exemple :</span><span class="sxs-lookup"><span data-stu-id="b620f-160">For example:</span></span> `function f(text?: string)`
* <span data-ttu-id="b620f-161">Donne une valeur par défaut au paramètre.</span><span class="sxs-lookup"><span data-stu-id="b620f-161">Give the parameter a default value.</span></span> <span data-ttu-id="b620f-162">Par exemple :</span><span class="sxs-lookup"><span data-stu-id="b620f-162">For example:</span></span> `function f(text: string = "abc")`

<span data-ttu-id="b620f-163">Pour une description détaillée du @paramvoir:[JSDoc](http://usejsdoc.org/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="b620f-163">For a detailed description of the code, see "HelloData Details."</span></span>

---
### <a name="requiresaddress"></a>@requiresAddress
<a id="requiresAddress"/>

<span data-ttu-id="b620f-164">Indique que l’adresse de la cellule dans laquelle la fonction est évaluée doit être fournie.</span><span class="sxs-lookup"><span data-stu-id="b620f-164">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span> 

<span data-ttu-id="b620f-165">Le dernier paramètre de la fonction doit être de type `CustomFunctions.Invocation` ou un type dérivé.</span><span class="sxs-lookup"><span data-stu-id="b620f-165">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="b620f-166">Lorsque la fonction est appelée, la`address` propriété contiendra l’adresse.</span><span class="sxs-lookup"><span data-stu-id="b620f-166">When the function is called, the `address` property will contain the address.</span></span>

---
### <a name="returns"></a>@returns
<a id="returns"/>

<span data-ttu-id="b620f-167">Syntaxe : @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="b620f-167">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="b620f-168">Fournit le type pour la valeur de retour.</span><span class="sxs-lookup"><span data-stu-id="b620f-168">Provides the type for the return value.</span></span>

<span data-ttu-id="b620f-169">Si `{type}` est omis, les informations de type TypeScript seront utilisées.</span><span class="sxs-lookup"><span data-stu-id="b620f-169">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="b620f-170">S’il n’existe aucune information type, le type sera `any`.</span><span class="sxs-lookup"><span data-stu-id="b620f-170">If there is no type info, the type will be `any`.</span></span>

---
### <a name="streaming"></a>@streaming
<a id="streaming"/>

<span data-ttu-id="b620f-171">Utilisé pour indiquer qu’une fonction personnalisée est une fonction diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="b620f-171">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="b620f-172">Le dernier paramètre doit être de type `CustomFunctions.StreamingInvocation<ResultType>`.</span><span class="sxs-lookup"><span data-stu-id="b620f-172">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="b620f-173">La fonction doit retourner `void`.</span><span class="sxs-lookup"><span data-stu-id="b620f-173">The function should return `void`.</span></span>

<span data-ttu-id="b620f-174">Les fonctions de diffusion en continu ne retournent pas de valeurs directement, mais doivent plutôt appeler `setResult(result: ResultType)` en utilisant le dernier paramètre.</span><span class="sxs-lookup"><span data-stu-id="b620f-174">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="b620f-175">Les exceptions levées par une fonction en continu sont ignorées.</span><span class="sxs-lookup"><span data-stu-id="b620f-175">Exceptions thrown by a streaming function are ignored.</span></span> `setResult()` <span data-ttu-id="b620f-176">peut être appelée avec l’erreur pour indiquer un résultat de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="b620f-176">may be called with Error to indicate an error result.</span></span>

<span data-ttu-id="b620f-177">Les fonctions en continu ne peuvent pas être marquées comme étant [ @volatile ](#volatile).</span><span class="sxs-lookup"><span data-stu-id="b620f-177">Streaming functions cannot be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a>@volatile
<a id="volatile"/>

<span data-ttu-id="b620f-178">Une fonction volatile est une dont le résultat ne peut pas être considéré comme le même à partir d’un moment à l’autre même si elle ne prend aucun argument ou les arguments n’ont pas changé.</span><span class="sxs-lookup"><span data-stu-id="b620f-178">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="b620f-179">Excel réévalue les cellules contenant des fonctions volatiles, ainsi que toutes les cellules dépendantes, chaque fois qu’il effectue un recalcul.</span><span class="sxs-lookup"><span data-stu-id="b620f-179">Excel reevaluates cells that contain volatile functions, together with all dependents, every time that it recalculates.</span></span> <span data-ttu-id="b620f-180">C’est pourquoi trop de dépendance des fonctions volatiles peut ralentir le recalcul.</span><span class="sxs-lookup"><span data-stu-id="b620f-180">For this reason, too much reliance on volatile functions can make recalculation times slow.</span></span>

<span data-ttu-id="b620f-181">Les fonctions en continu ne peuvent pas être volatiles.</span><span class="sxs-lookup"><span data-stu-id="b620f-181">Streaming functions cannot be volatile.</span></span>

---

## <a name="types"></a><span data-ttu-id="b620f-182">Types</span><span class="sxs-lookup"><span data-stu-id="b620f-182">Types</span></span>

<span data-ttu-id="b620f-183">En spécifiant un type de paramètre, Excel convertit les valeurs dans ce type avant d’appeler la fonction.</span><span class="sxs-lookup"><span data-stu-id="b620f-183">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="b620f-184">Si le type est`any`, aucune opération de conversion n’est effectuée.</span><span class="sxs-lookup"><span data-stu-id="b620f-184">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="b620f-185">Types de valeur</span><span class="sxs-lookup"><span data-stu-id="b620f-185">Value types</span></span>

<span data-ttu-id="b620f-186">Une valeur unique peut être représentée à l’aide d’un des types suivants : `boolean`, `number`, `string`.</span><span class="sxs-lookup"><span data-stu-id="b620f-186">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="b620f-187">Type matrice</span><span class="sxs-lookup"><span data-stu-id="b620f-187">Matrix type</span></span>

<span data-ttu-id="b620f-188">Utilisez une matrice à deux dimensions pour lesquels le paramètre ou la valeur de retour peut être une matrice de valeurs.</span><span class="sxs-lookup"><span data-stu-id="b620f-188">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="b620f-189">Par exemple, le type `number[][]` indique une matrice de nombres.</span><span class="sxs-lookup"><span data-stu-id="b620f-189">For example, the type `number[][]` indicates a matrix of numbers.</span></span> `string[][]` <span data-ttu-id="b620f-190">Indique une matrice de chaînes.</span><span class="sxs-lookup"><span data-stu-id="b620f-190">indicates a matrix of strings.</span></span> 

### <a name="error-type"></a><span data-ttu-id="b620f-191">Type d’erreur</span><span class="sxs-lookup"><span data-stu-id="b620f-191">Error Type</span></span>

<span data-ttu-id="b620f-192">Une fonction en non continu peut indiquer une erreur en retournant un type d’erreur.</span><span class="sxs-lookup"><span data-stu-id="b620f-192">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="b620f-193">Une fonction en continu peut indiquer une erreur en retournant un type d’erreur().</span><span class="sxs-lookup"><span data-stu-id="b620f-193">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="b620f-194">Promesse</span><span class="sxs-lookup"><span data-stu-id="b620f-194">Promise object.</span></span>

<span data-ttu-id="b620f-195">Une fonction peut renvoyer une promesse qui fournit la valeur lorsque la promesse aura été résolue.</span><span class="sxs-lookup"><span data-stu-id="b620f-195">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="b620f-196">Si la promesse est refusée, alors elle est considérée comme une erreur.</span><span class="sxs-lookup"><span data-stu-id="b620f-196">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="b620f-197">Autres types</span><span class="sxs-lookup"><span data-stu-id="b620f-197">Other solution types</span></span>

<span data-ttu-id="b620f-198">Un autre type sera traité comme une erreur.</span><span class="sxs-lookup"><span data-stu-id="b620f-198">Any other type will be treated as an error.</span></span>

## <a name="see-also"></a><span data-ttu-id="b620f-199">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b620f-199">See also</span></span>

* [<span data-ttu-id="b620f-200">Métadonnées des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b620f-200">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="b620f-201">Runtime pour les fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="b620f-201">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="b620f-202">Meilleures pratiques des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b620f-202">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="b620f-203">Fonctions personnalisées changelog</span><span class="sxs-lookup"><span data-stu-id="b620f-203">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="b620f-204">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="b620f-204">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="b620f-205">Débogage des métadonnées des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b620f-205">Custom functions debugging</span></span>](custom-functions-debugging.md)
