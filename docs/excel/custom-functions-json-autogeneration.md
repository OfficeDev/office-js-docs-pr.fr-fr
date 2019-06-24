---
ms.date: 06/18/2019
description: Utiliser les balises JSDOC pour créer dynamiquement vos fonctions personnalisées de métadonnées JSON.
title: Générer automatiquement des métadonnées JSON pour des fonctions personnalisées
localization_priority: Priority
ms.openlocfilehash: a02ca5fd67f29e1997579385e04d045f01e63bdb
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127904"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a><span data-ttu-id="cb84f-103">Générer automatiquement des métadonnées JSON pour des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="cb84f-103">Autogenerate JSON metadata for custom functions</span></span>

<span data-ttu-id="cb84f-104">Si vous écrivez une fonction Excel personnalisée en JavaScript ou TypeScript, vous pouvez utiliser les balises JSDoc pour la détailler en ajoutant des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="cb84f-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="cb84f-105">Les balises JSDoc sont ensuite utilisées lors de la génération pour créer le [fichier de métadonnées JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="cb84f-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="cb84f-106">En utilisant des balises JSDoc, vous n’avez plus besoin de modifier manuellement le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="cb84f-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="cb84f-107">Ajoutez la balise `@customfunction` dans les commentaires du code d’une fonction JavaScript ou TypeScript pour indiquer qu’il s’agit d’une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="cb84f-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="cb84f-108">Vous pouvez fournir les types de paramètres de la fonction en utilisant la balise[@param](#param)dans JavaScript, ou en précisant le [type de fonction](https://www.typescriptlang.org/docs/handbook/functions.html) dans TypeScript.</span><span class="sxs-lookup"><span data-stu-id="cb84f-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="cb84f-109">Si vous voulez en savoir plus, veuillez consulter les sections relatives à la balise[@param](#param) et aux sections[types](#types).</span><span class="sxs-lookup"><span data-stu-id="cb84f-109">For more information, see the [@param](#param) tag and [Types](#types) section.</span></span>

### <a name="adding-a-description-to-a-function"></a><span data-ttu-id="cb84f-110">Ajout d’une description à une fonction</span><span class="sxs-lookup"><span data-stu-id="cb84f-110">Adding a description to a function</span></span>

<span data-ttu-id="cb84f-111">La description s’affiche pour l’utilisateur sous forme de texte d’aide lorsqu’il a besoin d’aide pour comprendre le rôle de votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="cb84f-111">The description is displayed to the user as help text when they need help to understand what your custom function does.</span></span> <span data-ttu-id="cb84f-112">La description ne nécessite aucune balise spécifique.</span><span class="sxs-lookup"><span data-stu-id="cb84f-112">The description doesn't require any specific tag.</span></span> <span data-ttu-id="cb84f-113">Il vous suffit d’entrer une brève description dans le commentaire JSDoc.</span><span class="sxs-lookup"><span data-stu-id="cb84f-113">Just enter a short text description in the JSDoc comment.</span></span> <span data-ttu-id="cb84f-114">En général, la description est placée au début de la section commentaires JSDoc, mais elle fonctionnera peu importe son emplacement.</span><span class="sxs-lookup"><span data-stu-id="cb84f-114">In general the description is placed at the start of the JSDoc comment section, but it will work no matter where it is placed.</span></span>

<span data-ttu-id="cb84f-115">Pour consulter des exemples de descriptions de fonction intégrées, ouvrez Excel, accédez à l’onglet**formules** , puis sélectionnez **insérer une fonction**.</span><span class="sxs-lookup"><span data-stu-id="cb84f-115">To see examples of the built-in function descriptions, open Excel, go to the **Formulas** tab, and choose **Insert function**.</span></span> <span data-ttu-id="cb84f-116">Vous pouvez ensuite parcourir toutes les descriptions de fonction et voir vos propres fonctions personnalisées répertoriées.</span><span class="sxs-lookup"><span data-stu-id="cb84f-116">You can then browse through all the function descriptions, and also see your own custom functions listed.</span></span>

<span data-ttu-id="cb84f-117">Dans cet exemple, la phrase «calcule le volume d’une sphère.»</span><span class="sxs-lookup"><span data-stu-id="cb84f-117">In the following example, the phrase "Calculates the volume of a sphere."</span></span> <span data-ttu-id="cb84f-118">est la description de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="cb84f-118">is the description for the custom function.</span></span>

```JS
/**
/* Calculates the volume of a sphere
/* @customfunction VOLUME
...
 */
```


## <a name="jsdoc-tags"></a><span data-ttu-id="cb84f-119">Balises JSDoc</span><span class="sxs-lookup"><span data-stu-id="cb84f-119">JSDoc Tags</span></span>
<span data-ttu-id="cb84f-120">Voici quelles sont les balises JSDoc prises en charge dans les fonctions Excel personnalisées :</span><span class="sxs-lookup"><span data-stu-id="cb84f-120">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [<span data-ttu-id="cb84f-121">@ annulable</span><span class="sxs-lookup"><span data-stu-id="cb84f-121">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="cb84f-122">[@fonctionpersonnalisée](#customfunction)nom id</span><span class="sxs-lookup"><span data-stu-id="cb84f-122">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="cb84f-123">url[@urlaide](#helpurl)</span><span class="sxs-lookup"><span data-stu-id="cb84f-123">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="cb84f-124">[@param](#param) _{type}_ description nom</span><span class="sxs-lookup"><span data-stu-id="cb84f-124">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="cb84f-125">@requièreuneadresse</span><span class="sxs-lookup"><span data-stu-id="cb84f-125">@requiresAddress</span></span>](#requiresAddress)
* <span data-ttu-id="cb84f-126">[@renvoie](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="cb84f-126">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="cb84f-127">@diffusionencontinu</span><span class="sxs-lookup"><span data-stu-id="cb84f-127">@streaming</span></span>](#streaming)
* [<span data-ttu-id="cb84f-128">@volatile</span><span class="sxs-lookup"><span data-stu-id="cb84f-128">@volatile</span></span>](#volatile)

---
### <a name="cancelable"></a><span data-ttu-id="cb84f-129">@ annulable</span><span class="sxs-lookup"><span data-stu-id="cb84f-129">@cancelable</span></span>
<a id="cancelable"/>

<span data-ttu-id="cb84f-130">Indique qu’une fonction personnalisée souhaite effectuer une action lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="cb84f-130">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="cb84f-131">Le dernier paramètre de la fonction doit être de type `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="cb84f-131">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="cb84f-132">La fonction peut attribuer une fonction à la propriété `oncanceled` pour désigner l’action à effectuer lors de l’annulation de la fonction.</span><span class="sxs-lookup"><span data-stu-id="cb84f-132">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="cb84f-133">Si le dernier paramètre de fonction est de type `CustomFunctions.CancelableInvocation`, il sera considéré comme `@cancelable`, même si la balise n’apparaît pas.</span><span class="sxs-lookup"><span data-stu-id="cb84f-133">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="cb84f-134">Une fonction ne peut pas contenir les deux balises `@cancelable` et `@streaming`.</span><span class="sxs-lookup"><span data-stu-id="cb84f-134">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a><span data-ttu-id="cb84f-135">@fonctionpersonnalisée</span><span class="sxs-lookup"><span data-stu-id="cb84f-135">@customfunction</span></span>
<a id="customfunction"/>

<span data-ttu-id="cb84f-136">Syntaxe: @fonctionpersonnalisée_id_ _nom_</span><span class="sxs-lookup"><span data-stu-id="cb84f-136">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="cb84f-137">Spécifiez cette balise pour traiter la fonction JavaScript/TypeScript comme une fonction Excel personnalisée.</span><span class="sxs-lookup"><span data-stu-id="cb84f-137">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="cb84f-138">Cette balise est requise pour créer des métadonnées pour la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="cb84f-138">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="cb84f-139">Vous devez également insérer un appel vers`CustomFunctions.associate("id", functionName);`</span><span class="sxs-lookup"><span data-stu-id="cb84f-139">There should also be a call to `CustomFunctions.associate("id", functionName);`</span></span>

#### <a name="id"></a><span data-ttu-id="cb84f-140">id</span><span class="sxs-lookup"><span data-stu-id="cb84f-140">id</span></span>

<span data-ttu-id="cb84f-141">`id` Est un identificateur invariant pour la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="cb84f-141">The id is used as the invariant identifier for the custom function stored in the document.</span></span>

* <span data-ttu-id="cb84f-142">Si`id`n’est pas fourni, le nom de la fonction JavaScript/TypeScript est converti en majuscules, et les caractères rejetés sont supprimés.</span><span class="sxs-lookup"><span data-stu-id="cb84f-142">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="cb84f-143">Le `id`doit être unique pour toutes les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="cb84f-143">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="cb84f-144">Les caractères autorisés sont les suivants : A-Z, a-z, 0-9, traits de soulignement (\_) et point (.).</span><span class="sxs-lookup"><span data-stu-id="cb84f-144">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

#### <a name="name"></a><span data-ttu-id="cb84f-145">name</span><span class="sxs-lookup"><span data-stu-id="cb84f-145">name</span></span>

<span data-ttu-id="cb84f-146">Fournit le nom d’affichage `name`de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="cb84f-146">Provides the display name for the custom function.</span></span>

* <span data-ttu-id="cb84f-147">Si aucun nom n’est fourni, l’id servira aussi de nom.</span><span class="sxs-lookup"><span data-stu-id="cb84f-147">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="cb84f-148">Caractères autorisés : [caractères alphanumériques Unicode](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic) (lettres, chiffres), point (.) et trait de soulignement (\_).</span><span class="sxs-lookup"><span data-stu-id="cb84f-148">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="cb84f-149">Doit commencer par une lettre.</span><span class="sxs-lookup"><span data-stu-id="cb84f-149">Must start with a letter.</span></span>
* <span data-ttu-id="cb84f-150">Sa longueur maximale est limitée à 128 caractères.</span><span class="sxs-lookup"><span data-stu-id="cb84f-150">Maximum length is 128 characters.</span></span>

### <a name="description"></a><span data-ttu-id="cb84f-151">description</span><span class="sxs-lookup"><span data-stu-id="cb84f-151">description</span></span>

<span data-ttu-id="cb84f-152">Une description ne nécessite aucune balise spécifique.</span><span class="sxs-lookup"><span data-stu-id="cb84f-152">A description doesn't require any specific tag.</span></span> <span data-ttu-id="cb84f-153">Ajoutez une description à une fonction personnalisée en ajoutant une expression pour décrire le rôle de la fonction dans le commentaire JSDoc.</span><span class="sxs-lookup"><span data-stu-id="cb84f-153">Add a description to a custom function by adding a phrase to describe what the function does inside the JSDoc comment.</span></span> <span data-ttu-id="cb84f-154">Par défaut, le texte non balisé dans la section commentaire JSDoc est la description de la fonction.</span><span class="sxs-lookup"><span data-stu-id="cb84f-154">By default, whatever text is untagged in the JSDoc comment section will be the description of the function.</span></span> <span data-ttu-id="cb84f-155">La description s’affiche pour les utilisateurs dans Excel lors de la saisie de la fonction.</span><span class="sxs-lookup"><span data-stu-id="cb84f-155">The description appears to users in Excel as they are entering the function.</span></span> <span data-ttu-id="cb84f-156">Dans l’exemple suivant, l’expression «fonction qui calcule la somme de deux nombres» est la description de la fonction personnalisée dont la propriété ID est`SUM`.</span><span class="sxs-lookup"><span data-stu-id="cb84f-156">In the following example, the phrase "A function that sums two numbers" is the description for the custom function with the id property of `SUM`.</span></span>

```JS
/**
/* @customfunction SUM
/* A function that sums two numbers
...
 */
```

---
### <a name="helpurl"></a><span data-ttu-id="cb84f-157">@urlaide</span><span class="sxs-lookup"><span data-stu-id="cb84f-157">@helpurl</span></span>
<a id="helpurl"/>

<span data-ttu-id="cb84f-158">Syntaxe: @urlaide_url_</span><span class="sxs-lookup"><span data-stu-id="cb84f-158">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="cb84f-159">L’_url_ fournie est affichée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="cb84f-159">The provided _url_ is displayed in Excel.</span></span>

---
### <a name="param"></a><span data-ttu-id="cb84f-160">@param</span><span class="sxs-lookup"><span data-stu-id="cb84f-160">@param</span></span>
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="cb84f-161">JavaScript</span><span class="sxs-lookup"><span data-stu-id="cb84f-161">JavaScript</span></span>

<span data-ttu-id="cb84f-162">Syntaxe JavaScript : @param {type} nom_description_</span><span class="sxs-lookup"><span data-stu-id="cb84f-162">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="cb84f-163">`{type}`doit spécifier les informations de type entre deux accolades.</span><span class="sxs-lookup"><span data-stu-id="cb84f-163">`{type}` should specify the type info within curly braces.</span></span> <span data-ttu-id="cb84f-164">Consultez la section [Types](##types) pour savoir quels types peuvent être utilisés.</span><span class="sxs-lookup"><span data-stu-id="cb84f-164">See the [Types](##types) for more information about the types which may be used.</span></span> <span data-ttu-id="cb84f-165">Facultatif : si aucun serveur n’est spécifié, le type `any` sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="cb84f-165">Optional: if not specified, the type `any` will be used.</span></span>
* <span data-ttu-id="cb84f-166">`name`spécifie le paramètre auquel s’applique la balise.</span><span class="sxs-lookup"><span data-stu-id="cb84f-166">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="cb84f-167">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="cb84f-167">Required.</span></span>
* <span data-ttu-id="cb84f-168">`description`fournit la description qui s’affiche dans Excel pour le paramètre de la fonction.</span><span class="sxs-lookup"><span data-stu-id="cb84f-168">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="cb84f-169">Facultatif.</span><span class="sxs-lookup"><span data-stu-id="cb84f-169">Optional.</span></span>

<span data-ttu-id="cb84f-170">Pour désigner un paramètre de fonction personnalisée comme étant facultatif :</span><span class="sxs-lookup"><span data-stu-id="cb84f-170">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="cb84f-171">Placez les crochets autour du nom du paramètre.</span><span class="sxs-lookup"><span data-stu-id="cb84f-171">Put square brackets around the parameter name.</span></span> <span data-ttu-id="cb84f-172">Par exemple : `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="cb84f-172">For example: `@param {string} [text] Optional text`.</span></span>

> [!NOTE]
> <span data-ttu-id="cb84f-173">La valeur par défaut pour les paramètres facultatifs est `null`.</span><span class="sxs-lookup"><span data-stu-id="cb84f-173">The default value for optional parameters is `null`.</span></span>

#### <a name="typescript"></a><span data-ttu-id="cb84f-174">TypeScript</span><span class="sxs-lookup"><span data-stu-id="cb84f-174">TypeScript</span></span>

<span data-ttu-id="cb84f-175">Syntaxe TypeScript : nom @param_description_</span><span class="sxs-lookup"><span data-stu-id="cb84f-175">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="cb84f-176">`name`spécifie le paramètre auquel s’applique la balise.</span><span class="sxs-lookup"><span data-stu-id="cb84f-176">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="cb84f-177">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="cb84f-177">Required.</span></span>
* <span data-ttu-id="cb84f-178">`description`fournit la description qui s’affiche dans Excel pour le paramètre de la fonction.</span><span class="sxs-lookup"><span data-stu-id="cb84f-178">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="cb84f-179">Facultatif.</span><span class="sxs-lookup"><span data-stu-id="cb84f-179">Optional.</span></span>

<span data-ttu-id="cb84f-180">Consultez la section [Types](##types) pour savoir quels types de paramètres de fonction peuvent être utilisés.</span><span class="sxs-lookup"><span data-stu-id="cb84f-180">See the [Types](##types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="cb84f-181">Pour désigner un paramètre de fonction personnalisée comme étant facultatif, effectuez l’une des actions suivantes :</span><span class="sxs-lookup"><span data-stu-id="cb84f-181">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="cb84f-182">Utilisez un paramètre facultatif.</span><span class="sxs-lookup"><span data-stu-id="cb84f-182">Use an optional parameter.</span></span> <span data-ttu-id="cb84f-183">Par exemple : `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="cb84f-183">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="cb84f-184">Définissez ce paramètre sur une valeur par défaut.</span><span class="sxs-lookup"><span data-stu-id="cb84f-184">Give the parameter a default value.</span></span> <span data-ttu-id="cb84f-185">Par exemple : `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="cb84f-185">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="cb84f-186">Pour consulter une description détaillée du @param, reportez-vous à la page suivante : [JSDoc](https://jsdoc.app/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="cb84f-186">For detailed description of the @param see: [JSDoc](https://jsdoc.app/tags-param.html)</span></span>

> [!NOTE]
> <span data-ttu-id="cb84f-187">La valeur par défaut pour les paramètres facultatifs est `null`.</span><span class="sxs-lookup"><span data-stu-id="cb84f-187">The default value for optional parameters is `null`.</span></span>

---
### <a name="requiresaddress"></a><span data-ttu-id="cb84f-188">@requièreuneadresse</span><span class="sxs-lookup"><span data-stu-id="cb84f-188">@requiresAddress</span></span>
<a id="requiresAddress"/>

<span data-ttu-id="cb84f-189">Indique que l’adresse de la cellule dans laquelle la fonction est évaluée doit être fournie.</span><span class="sxs-lookup"><span data-stu-id="cb84f-189">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span> 

<span data-ttu-id="cb84f-190">Le dernier paramètre de la fonction doit être de type `CustomFunctions.Invocation` ou un type dérivé.</span><span class="sxs-lookup"><span data-stu-id="cb84f-190">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="cb84f-191">Lorsque la fonction est appelée, la propriété `address` contiendra l’adresse.</span><span class="sxs-lookup"><span data-stu-id="cb84f-191">When the function is called, the `address` property will contain the address.</span></span>

---
### <a name="returns"></a><span data-ttu-id="cb84f-192">@renvoie :</span><span class="sxs-lookup"><span data-stu-id="cb84f-192">@returns</span></span>
<a id="returns"/>

<span data-ttu-id="cb84f-193">Syntaxe: @renvoie {_type_}</span><span class="sxs-lookup"><span data-stu-id="cb84f-193">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="cb84f-194">Fournit le type pour la valeur renvoyée.</span><span class="sxs-lookup"><span data-stu-id="cb84f-194">Provides the type for the return value.</span></span>

<span data-ttu-id="cb84f-195">Si `{type}` est omis, les informations de type TypeScript seront utilisées.</span><span class="sxs-lookup"><span data-stu-id="cb84f-195">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="cb84f-196">S’il n’existe aucune information définissant le type, ce dernier sera `any`.</span><span class="sxs-lookup"><span data-stu-id="cb84f-196">If there is no type info, the type will be `any`.</span></span>

---
### <a name="streaming"></a><span data-ttu-id="cb84f-197">@diffusionencontinu</span><span class="sxs-lookup"><span data-stu-id="cb84f-197">@streaming</span></span>
<a id="streaming"/>

<span data-ttu-id="cb84f-198">Utilisé pour indiquer qu’une fonction personnalisée est une fonction diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="cb84f-198">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="cb84f-199">Le dernier paramètre doit être de type `CustomFunctions.StreamingInvocation<ResultType>`.</span><span class="sxs-lookup"><span data-stu-id="cb84f-199">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="cb84f-200">La fonction doit renvoyer `void`.</span><span class="sxs-lookup"><span data-stu-id="cb84f-200">The function should return `void`.</span></span>

<span data-ttu-id="cb84f-201">Les fonctions de diffusion en continu ne renvoient pas de valeurs directement, mais doivent plutôt appeler `setResult(result: ResultType)` en utilisant le dernier paramètre.</span><span class="sxs-lookup"><span data-stu-id="cb84f-201">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="cb84f-202">Les exceptions levées par une fonction en continu sont ignorées.</span><span class="sxs-lookup"><span data-stu-id="cb84f-202">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="cb84f-203">`setResult()`peut être appelée avec Error pour indiquer un résultat erroné.</span><span class="sxs-lookup"><span data-stu-id="cb84f-203">`setResult()` may be called with Error to indicate an error result.</span></span>

<span data-ttu-id="cb84f-204">Vous ne pouvez pas utiliser les balises en diffusion en continu comme [@volatile](#volatile).</span><span class="sxs-lookup"><span data-stu-id="cb84f-204">Streaming functions cannot be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a><span data-ttu-id="cb84f-205">@volatile</span><span class="sxs-lookup"><span data-stu-id="cb84f-205">@volatile</span></span>
<a id="volatile"/>

<span data-ttu-id="cb84f-206">Une fonction volatile est une fonction dont le résultat peut changer d’un moment à l’autre, même si elle ne récupère pas d’argument ou si ses arguments ne changent pas.</span><span class="sxs-lookup"><span data-stu-id="cb84f-206">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="cb84f-207">À chaque calcul, Excel réévalue les cellules contenant des fonctions volatiles, ainsi que toutes leurs cellules dépendantes.</span><span class="sxs-lookup"><span data-stu-id="cb84f-207">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="cb84f-208">C’est pourquoi, un trop grand nombre de dépendances de fonctions volatiles risque de ralentir les calculs. Nous vous recommandons d’en utiliser aussi peu que possible.</span><span class="sxs-lookup"><span data-stu-id="cb84f-208">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="cb84f-209">Les fonctions de diffusion en continu ne peuvent pas être volatiles.</span><span class="sxs-lookup"><span data-stu-id="cb84f-209">Streaming functions cannot be volatile.</span></span>

---

## <a name="types"></a><span data-ttu-id="cb84f-210">Types</span><span class="sxs-lookup"><span data-stu-id="cb84f-210">Types</span></span>

<span data-ttu-id="cb84f-211">En spécifiant un type de paramètre, Excel convertit les valeurs en ce type, avant d’appeler la fonction.</span><span class="sxs-lookup"><span data-stu-id="cb84f-211">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="cb84f-212">Si le type est `any`, Excel n’effectue pas de conversion.</span><span class="sxs-lookup"><span data-stu-id="cb84f-212">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="cb84f-213">Types de valeur</span><span class="sxs-lookup"><span data-stu-id="cb84f-213">Value types</span></span>

<span data-ttu-id="cb84f-214">Une valeur unique peut être représentée à l’aide d’un des types suivants : `boolean`, `number` ou `string`.</span><span class="sxs-lookup"><span data-stu-id="cb84f-214">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="cb84f-215">Type matrice</span><span class="sxs-lookup"><span data-stu-id="cb84f-215">Matrix type</span></span>

<span data-ttu-id="cb84f-216">Utilisez une matrice à deux dimensions pour que le paramètre ou la valeur renvoyée soit une matrice de valeurs.</span><span class="sxs-lookup"><span data-stu-id="cb84f-216">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="cb84f-217">Par exemple, le type `number[][]` indique une matrice de nombres.</span><span class="sxs-lookup"><span data-stu-id="cb84f-217">For example, the type `number[][]` indicates a matrix of numbers.</span></span> <span data-ttu-id="cb84f-218">`string[][]`indique une matrice de chaînes.</span><span class="sxs-lookup"><span data-stu-id="cb84f-218">`string[][]` indicates a matrix of strings.</span></span> 

### <a name="error-type"></a><span data-ttu-id="cb84f-219">Type d’erreur</span><span class="sxs-lookup"><span data-stu-id="cb84f-219">Error type</span></span>

<span data-ttu-id="cb84f-220">Une fonction qui n’est pas une fonction de diffusion en continu peut indiquer une erreur en renvoyant un type Error.</span><span class="sxs-lookup"><span data-stu-id="cb84f-220">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="cb84f-221">Une fonction de diffusion en continu peut indiquer une erreur en appelant`setResult()`avec un type Error.</span><span class="sxs-lookup"><span data-stu-id="cb84f-221">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="cb84f-222">Promise</span><span class="sxs-lookup"><span data-stu-id="cb84f-222">Promise</span></span>

<span data-ttu-id="cb84f-223">Une fonction peut renvoyer un objet Promise (pour « promesse »). Ce dernier fournit une valeur lors de sa résolution.</span><span class="sxs-lookup"><span data-stu-id="cb84f-223">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="cb84f-224">Si la résolution de l’objet Promise est refusée, cela entraîne une erreur.</span><span class="sxs-lookup"><span data-stu-id="cb84f-224">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="cb84f-225">Autres types</span><span class="sxs-lookup"><span data-stu-id="cb84f-225">Other types</span></span>

<span data-ttu-id="cb84f-226">Tout autre type sera traité comme une erreur.</span><span class="sxs-lookup"><span data-stu-id="cb84f-226">Any other type will be treated as an error.</span></span>

## <a name="next-steps"></a><span data-ttu-id="cb84f-227">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="cb84f-227">Next steps</span></span>
<span data-ttu-id="cb84f-228">Découvrez les [conventions d’affectation des noms des fonctions personnalisées](custom-functions-naming.md).</span><span class="sxs-lookup"><span data-stu-id="cb84f-228">Learn about [naming conventions for custom functions](custom-functions-naming.md).</span></span> <span data-ttu-id="cb84f-229">Découvrez également comment [localiser vos fonctions](custom-functions-localize.md), ce qui implique que vous [écriviez votre fichier JSON à la main](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="cb84f-229">Alternatively, learn how to [localize your functions](custom-functions-localize.md) which requires you to [write your JSON file by hand](custom-functions-json.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="cb84f-230">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="cb84f-230">See also</span></span>

* [<span data-ttu-id="cb84f-231">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="cb84f-231">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="cb84f-232">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="cb84f-232">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="cb84f-233">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="cb84f-233">Create custom functions in Excel</span></span>](custom-functions-overview.md)
