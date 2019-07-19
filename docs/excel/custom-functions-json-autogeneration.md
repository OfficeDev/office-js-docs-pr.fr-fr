---
ms.date: 07/15/2019
description: Utiliser les balises JSDOC pour créer dynamiquement vos fonctions personnalisées de métadonnées JSON.
title: Générer automatiquement des métadonnées JSON pour des fonctions personnalisées
localization_priority: Priority
ms.openlocfilehash: afcfb6ff869acf1d508ebda7fc242dd9724bf165
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771315"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a><span data-ttu-id="b1013-103">Générer automatiquement des métadonnées JSON pour des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b1013-103">Autogenerate JSON metadata for custom functions</span></span>

<span data-ttu-id="b1013-104">Si vous écrivez une fonction Excel personnalisée en JavaScript ou TypeScript, vous pouvez utiliser les [balises JSDoc](https://jsdoc.app/) pour la détailler en ajoutant des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="b1013-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="b1013-105">Les balises JSDoc sont ensuite utilisées lors de la génération pour créer le [fichier de métadonnées JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="b1013-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="b1013-106">En utilisant des balises JSDoc, vous n’avez plus besoin de modifier manuellement le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="b1013-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="b1013-107">Ajoutez la balise `@customfunction` dans les commentaires du code d’une fonction JavaScript ou TypeScript pour indiquer qu’il s’agit d’une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="b1013-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="b1013-108">Vous pouvez fournir les types de paramètres de la fonction en utilisant la balise[@param](#param)dans JavaScript, ou en précisant le [type de fonction](https://www.typescriptlang.org/docs/handbook/functions.html) dans TypeScript.</span><span class="sxs-lookup"><span data-stu-id="b1013-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="b1013-109">Si vous voulez en savoir plus, veuillez consulter les sections relatives à la balise[@param](#param) et aux sections[types](#types).</span><span class="sxs-lookup"><span data-stu-id="b1013-109">For more information, see the [@param](#param) tag and [Types](#types) section.</span></span>

### <a name="adding-a-description-to-a-function"></a><span data-ttu-id="b1013-110">Ajout d’une description à une fonction</span><span class="sxs-lookup"><span data-stu-id="b1013-110">Adding a description to a function</span></span>

<span data-ttu-id="b1013-111">La description s’affiche pour l’utilisateur sous forme de texte d’aide lorsqu’il a besoin d’aide pour comprendre le rôle de votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="b1013-111">The description is displayed to the user as help text when they need help to understand what your custom function does.</span></span> <span data-ttu-id="b1013-112">La description ne nécessite aucune balise spécifique.</span><span class="sxs-lookup"><span data-stu-id="b1013-112">The description doesn't require any specific tag.</span></span> <span data-ttu-id="b1013-113">Il vous suffit d’entrer une brève description dans le commentaire JSDoc.</span><span class="sxs-lookup"><span data-stu-id="b1013-113">Just enter a short text description in the JSDoc comment.</span></span> <span data-ttu-id="b1013-114">En général, la description est placée au début de la section commentaires JSDoc, mais elle fonctionnera peu importe son emplacement.</span><span class="sxs-lookup"><span data-stu-id="b1013-114">In general the description is placed at the start of the JSDoc comment section, but it will work no matter where it is placed.</span></span>

<span data-ttu-id="b1013-115">Pour consulter des exemples de descriptions de fonction intégrées, ouvrez Excel, accédez à l’onglet**formules** , puis sélectionnez **insérer une fonction**.</span><span class="sxs-lookup"><span data-stu-id="b1013-115">To see examples of the built-in function descriptions, open Excel, go to the **Formulas** tab, and choose **Insert function**.</span></span> <span data-ttu-id="b1013-116">Vous pouvez ensuite parcourir toutes les descriptions de fonction et voir vos propres fonctions personnalisées répertoriées.</span><span class="sxs-lookup"><span data-stu-id="b1013-116">You can then browse through all the function descriptions, and also see your own custom functions listed.</span></span>

<span data-ttu-id="b1013-117">Dans cet exemple, la phrase «calcule le volume d’une sphère.»</span><span class="sxs-lookup"><span data-stu-id="b1013-117">In the following example, the phrase "Calculates the volume of a sphere."</span></span> <span data-ttu-id="b1013-118">est la description de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="b1013-118">is the description for the custom function.</span></span>

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```


## <a name="jsdoc-tags"></a><span data-ttu-id="b1013-119">Balises JSDoc</span><span class="sxs-lookup"><span data-stu-id="b1013-119">JSDoc Tags</span></span>
<span data-ttu-id="b1013-120">Voici quelles sont les balises JSDoc prises en charge dans les fonctions Excel personnalisées :</span><span class="sxs-lookup"><span data-stu-id="b1013-120">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [<span data-ttu-id="b1013-121">@ annulable</span><span class="sxs-lookup"><span data-stu-id="b1013-121">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="b1013-122">[@fonctionpersonnalisée](#customfunction)nom id</span><span class="sxs-lookup"><span data-stu-id="b1013-122">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="b1013-123">url[@urlaide](#helpurl)</span><span class="sxs-lookup"><span data-stu-id="b1013-123">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="b1013-124">[@param](#param) _{type}_ description nom</span><span class="sxs-lookup"><span data-stu-id="b1013-124">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="b1013-125">@requièreuneadresse</span><span class="sxs-lookup"><span data-stu-id="b1013-125">@requiresAddress</span></span>](#requiresAddress)
* <span data-ttu-id="b1013-126">[@renvoie](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="b1013-126">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="b1013-127">@diffusionencontinu</span><span class="sxs-lookup"><span data-stu-id="b1013-127">@streaming</span></span>](#streaming)
* [<span data-ttu-id="b1013-128">@volatile</span><span class="sxs-lookup"><span data-stu-id="b1013-128">@volatile</span></span>](#volatile)

---
### <a name="cancelable"></a><span data-ttu-id="b1013-129">@ annulable</span><span class="sxs-lookup"><span data-stu-id="b1013-129">@cancelable</span></span>
<a id="cancelable"/>

<span data-ttu-id="b1013-130">Indique qu’une fonction personnalisée souhaite effectuer une action lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="b1013-130">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="b1013-131">Le dernier paramètre de la fonction doit être de type `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="b1013-131">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="b1013-132">La fonction peut attribuer une fonction à la propriété `oncanceled` pour désigner l’action à effectuer lors de l’annulation de la fonction.</span><span class="sxs-lookup"><span data-stu-id="b1013-132">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="b1013-133">Si le dernier paramètre de fonction est de type `CustomFunctions.CancelableInvocation`, il sera considéré comme `@cancelable`, même si la balise n’apparaît pas.</span><span class="sxs-lookup"><span data-stu-id="b1013-133">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="b1013-134">Une fonction ne peut pas contenir les deux balises `@cancelable` et `@streaming`.</span><span class="sxs-lookup"><span data-stu-id="b1013-134">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a><span data-ttu-id="b1013-135">@fonctionpersonnalisée</span><span class="sxs-lookup"><span data-stu-id="b1013-135">@customfunction</span></span>
<a id="customfunction"/>

<span data-ttu-id="b1013-136">Syntaxe: @fonctionpersonnalisée_id_ _nom_</span><span class="sxs-lookup"><span data-stu-id="b1013-136">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="b1013-137">Spécifiez cette balise pour traiter la fonction JavaScript/TypeScript comme une fonction Excel personnalisée.</span><span class="sxs-lookup"><span data-stu-id="b1013-137">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="b1013-138">Cette balise est requise pour créer des métadonnées pour la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="b1013-138">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="b1013-139">L’exemple suivant illustre la méthode la plus simple pour déclarer une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="b1013-139">The following example shows the simplest way to declare a custom function.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a><span data-ttu-id="b1013-140">id</span><span class="sxs-lookup"><span data-stu-id="b1013-140">id</span></span>

<span data-ttu-id="b1013-141">`id` Est un identificateur invariant pour la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="b1013-141">The id is used as the invariant identifier for the custom function stored in the document.</span></span>

* <span data-ttu-id="b1013-142">Si`id`n’est pas fourni, le nom de la fonction JavaScript/TypeScript est converti en majuscules, et les caractères rejetés sont supprimés.</span><span class="sxs-lookup"><span data-stu-id="b1013-142">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="b1013-143">Le `id`doit être unique pour toutes les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="b1013-143">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="b1013-144">Les caractères autorisés sont les suivants : A-Z, a-z, 0-9, traits de soulignement (\_) et point (.).</span><span class="sxs-lookup"><span data-stu-id="b1013-144">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

<span data-ttu-id="b1013-145">Dans l’exemple suivant, Increments correspond à l’`id` et au `name` de la fonction.</span><span class="sxs-lookup"><span data-stu-id="b1013-145">In the following example, increment is the `id` and the `name` of the function.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INCREMENT
 * ...
 */
```

#### <a name="name"></a><span data-ttu-id="b1013-146">name</span><span class="sxs-lookup"><span data-stu-id="b1013-146">name</span></span>

<span data-ttu-id="b1013-147">Fournit le nom d’affichage `name`de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="b1013-147">Provides the display name for the custom function.</span></span>

* <span data-ttu-id="b1013-148">Si aucun nom n’est fourni, l’id servira aussi de nom.</span><span class="sxs-lookup"><span data-stu-id="b1013-148">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="b1013-149">Caractères autorisés : [caractères alphanumériques Unicode](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic) (lettres, chiffres), point (.) et trait de soulignement (\_).</span><span class="sxs-lookup"><span data-stu-id="b1013-149">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="b1013-150">Doit commencer par une lettre.</span><span class="sxs-lookup"><span data-stu-id="b1013-150">Must start with a letter.</span></span>
* <span data-ttu-id="b1013-151">Sa longueur maximale est limitée à 128 caractères.</span><span class="sxs-lookup"><span data-stu-id="b1013-151">Maximum length is 128 characters.</span></span>

<span data-ttu-id="b1013-152">Dans l’exemple suivant, INC correspond à l’`id` de la fonction, tandis que `increment` correspond au `name`.</span><span class="sxs-lookup"><span data-stu-id="b1013-152">In the following example, INC is the `id` of the function and `increment` is the `name`.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a><span data-ttu-id="b1013-153">description</span><span class="sxs-lookup"><span data-stu-id="b1013-153">description</span></span>

<span data-ttu-id="b1013-154">Une description ne nécessite aucune balise spécifique.</span><span class="sxs-lookup"><span data-stu-id="b1013-154">A description doesn't require any specific tag.</span></span> <span data-ttu-id="b1013-155">Ajoutez une description à une fonction personnalisée en ajoutant une expression pour décrire le rôle de la fonction dans le commentaire JSDoc.</span><span class="sxs-lookup"><span data-stu-id="b1013-155">Add a description to a custom function by adding a phrase to describe what the function does inside the JSDoc comment.</span></span> <span data-ttu-id="b1013-156">Par défaut, le texte non balisé dans la section commentaire JSDoc est la description de la fonction.</span><span class="sxs-lookup"><span data-stu-id="b1013-156">By default, whatever text is untagged in the JSDoc comment section will be the description of the function.</span></span> <span data-ttu-id="b1013-157">La description s’affiche dans Excel lorsque l’utilisateur saisit la fonction.</span><span class="sxs-lookup"><span data-stu-id="b1013-157">The description appears to users in Excel as they are entering the function.</span></span> <span data-ttu-id="b1013-158">Dans l’exemple suivant, la phrase « A function that adds two numbers » (« Une fonction qui ajoute deux nombres ») est la description de la fonction personnalisée dont la propriété ID est `ADD`.</span><span class="sxs-lookup"><span data-stu-id="b1013-158">In the following example, the phrase "A function that adds two numbers" is the description for the custom function with the id property of `ADD`.</span></span>

<span data-ttu-id="b1013-159">Dans l’exemple suivant, ADD correspond à l’`id` et au `name` de la fonction. Une description est indiquée.</span><span class="sxs-lookup"><span data-stu-id="b1013-159">In the following example, ADD is the `id` and `name` of the function and a description is given.</span></span>

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

---
### <a name="helpurl"></a><span data-ttu-id="b1013-160">@urlaide</span><span class="sxs-lookup"><span data-stu-id="b1013-160">@helpurl</span></span>
<a id="helpurl"/>

<span data-ttu-id="b1013-161">Syntaxe: @urlaide_url_</span><span class="sxs-lookup"><span data-stu-id="b1013-161">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="b1013-162">L’_url_ fournie est affichée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="b1013-162">The provided _url_ is displayed in Excel.</span></span>

<span data-ttu-id="b1013-163">Dans l’exemple suivant, l’`helpurl` est www.contoso.com/weatherhelp.</span><span class="sxs-lookup"><span data-stu-id="b1013-163">In the following example, the `helpurl` is www.contoso.com/weatherhelp.</span></span>

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

---
### <a name="param"></a><span data-ttu-id="b1013-164">@param</span><span class="sxs-lookup"><span data-stu-id="b1013-164">@param</span></span>
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="b1013-165">JavaScript</span><span class="sxs-lookup"><span data-stu-id="b1013-165">JavaScript</span></span>

<span data-ttu-id="b1013-166">Syntaxe JavaScript : @param {type} nom_description_</span><span class="sxs-lookup"><span data-stu-id="b1013-166">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="b1013-167">`{type}`doit spécifier les informations de type entre deux accolades.</span><span class="sxs-lookup"><span data-stu-id="b1013-167">`{type}` should specify the type info within curly braces.</span></span> <span data-ttu-id="b1013-168">Consultez la section [Types](#types) pour savoir quels types peuvent être utilisés.</span><span class="sxs-lookup"><span data-stu-id="b1013-168">See the [Types](#types) for more information about the types which may be used.</span></span> <span data-ttu-id="b1013-169">Facultatif : si aucun serveur n’est spécifié, le type `any` sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="b1013-169">Optional: if not specified, the type `any` will be used.</span></span>
* <span data-ttu-id="b1013-170">`name`spécifie le paramètre auquel s’applique la balise.</span><span class="sxs-lookup"><span data-stu-id="b1013-170">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="b1013-171">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="b1013-171">Required.</span></span>
* <span data-ttu-id="b1013-172">`description`fournit la description qui s’affiche dans Excel pour le paramètre de la fonction.</span><span class="sxs-lookup"><span data-stu-id="b1013-172">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="b1013-173">Facultatif.</span><span class="sxs-lookup"><span data-stu-id="b1013-173">Optional.</span></span>

<span data-ttu-id="b1013-174">Pour désigner un paramètre de fonction personnalisée comme étant facultatif :</span><span class="sxs-lookup"><span data-stu-id="b1013-174">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="b1013-175">Placez les crochets autour du nom du paramètre.</span><span class="sxs-lookup"><span data-stu-id="b1013-175">Put square brackets around the parameter name.</span></span> <span data-ttu-id="b1013-176">Par exemple : `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="b1013-176">For example: `@param {string} [text] Optional text`.</span></span>

> [!NOTE]
> <span data-ttu-id="b1013-177">La valeur par défaut pour les paramètres facultatifs est `null`.</span><span class="sxs-lookup"><span data-stu-id="b1013-177">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="b1013-178">L’exemple suivant représente une fonction ADD qui ajoute deux ou trois nombres, où le troisième nombre est un paramètre facultatif.</span><span class="sxs-lookup"><span data-stu-id="b1013-178">The following example shows a ADD function which adds two or three numbers, with the third number as an optional parameter.</span></span>

```js
/**
 * A function which sums two, or optionally three, numbers.
 * @customfunction ADDNUMBERS
 * @param firstNumber {number} First number to add.
 * @param secondNumber {number} Second number to add.
 * @param [thirdNumber] {number} Optional third number you wish to add.
 * ...
 */
```

#### <a name="typescript"></a><span data-ttu-id="b1013-179">TypeScript</span><span class="sxs-lookup"><span data-stu-id="b1013-179">TypeScript</span></span>

<span data-ttu-id="b1013-180">Syntaxe TypeScript : nom @param_description_</span><span class="sxs-lookup"><span data-stu-id="b1013-180">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="b1013-181">`name`spécifie le paramètre auquel s’applique la balise.</span><span class="sxs-lookup"><span data-stu-id="b1013-181">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="b1013-182">Obligatoire.</span><span class="sxs-lookup"><span data-stu-id="b1013-182">Required.</span></span>
* <span data-ttu-id="b1013-183">`description`fournit la description qui s’affiche dans Excel pour le paramètre de la fonction.</span><span class="sxs-lookup"><span data-stu-id="b1013-183">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="b1013-184">Facultatif.</span><span class="sxs-lookup"><span data-stu-id="b1013-184">Optional.</span></span>

<span data-ttu-id="b1013-185">Consultez la section [Types](#types) pour savoir quels types de paramètres de fonction peuvent être utilisés.</span><span class="sxs-lookup"><span data-stu-id="b1013-185">See the [Types](#types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="b1013-186">Pour désigner un paramètre de fonction personnalisée comme étant facultatif, effectuez l’une des actions suivantes :</span><span class="sxs-lookup"><span data-stu-id="b1013-186">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="b1013-187">Utilisez un paramètre facultatif.</span><span class="sxs-lookup"><span data-stu-id="b1013-187">Use an optional parameter.</span></span> <span data-ttu-id="b1013-188">Par exemple : `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="b1013-188">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="b1013-189">Définissez ce paramètre sur une valeur par défaut.</span><span class="sxs-lookup"><span data-stu-id="b1013-189">Give the parameter a default value.</span></span> <span data-ttu-id="b1013-190">Par exemple : `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="b1013-190">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="b1013-191">Pour consulter une description détaillée du @param, reportez-vous à la page suivante : [JSDoc](https://jsdoc.app/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="b1013-191">For detailed description of the @param see: [JSDoc](https://jsdoc.app/tags-param.html)</span></span>

> [!NOTE]
> <span data-ttu-id="b1013-192">La valeur par défaut pour les paramètres facultatifs est `null`.</span><span class="sxs-lookup"><span data-stu-id="b1013-192">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="b1013-193">L’exemple suivant représente la fonction `add` qui ajoute deux nombres.</span><span class="sxs-lookup"><span data-stu-id="b1013-193">The following example shows the `add` function that adds two numbers.</span></span>

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

---
### <a name="requiresaddress"></a><span data-ttu-id="b1013-194">@requièreuneadresse</span><span class="sxs-lookup"><span data-stu-id="b1013-194">@requiresAddress</span></span>
<a id="requiresAddress"/>

<span data-ttu-id="b1013-195">Indique que l’adresse de la cellule dans laquelle la fonction est évaluée doit être fournie.</span><span class="sxs-lookup"><span data-stu-id="b1013-195">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span>

<span data-ttu-id="b1013-196">Le dernier paramètre de la fonction doit être de type `CustomFunctions.Invocation` ou un type dérivé.</span><span class="sxs-lookup"><span data-stu-id="b1013-196">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="b1013-197">Lorsque la fonction est appelée, la propriété `address` contiendra l’adresse.</span><span class="sxs-lookup"><span data-stu-id="b1013-197">When the function is called, the `address` property will contain the address.</span></span> <span data-ttu-id="b1013-198">Si vous souhaitez consulter un exemple de fonction utilisant la balise `@requiresAddress`, veuillez vous reporter à la section [Adressage du paramètre de contexte d’une cellule](custom-functions-parameter-options.md#addressing-cells-context-parameter).</span><span class="sxs-lookup"><span data-stu-id="b1013-198">For an example of a function that uses the `@requiresAddress` tag, see [Addressing cell's context parameter](custom-functions-parameter-options.md#addressing-cells-context-parameter).</span></span>

---
### <a name="returns"></a><span data-ttu-id="b1013-199">@renvoie :</span><span class="sxs-lookup"><span data-stu-id="b1013-199">@returns</span></span>
<a id="returns"/>

<span data-ttu-id="b1013-200">Syntaxe: @renvoie {_type_}</span><span class="sxs-lookup"><span data-stu-id="b1013-200">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="b1013-201">Fournit le type pour la valeur renvoyée.</span><span class="sxs-lookup"><span data-stu-id="b1013-201">Provides the type for the return value.</span></span>

<span data-ttu-id="b1013-202">Si `{type}` est omis, les informations de type TypeScript seront utilisées.</span><span class="sxs-lookup"><span data-stu-id="b1013-202">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="b1013-203">S’il n’existe aucune information définissant le type, ce dernier sera `any`.</span><span class="sxs-lookup"><span data-stu-id="b1013-203">If there is no type info, the type will be `any`.</span></span>

<span data-ttu-id="b1013-204">L’exemple suivant représente la fonction `add` qui utilise la balise `@returns`.</span><span class="sxs-lookup"><span data-stu-id="b1013-204">The following example shows the `add` function that uses the `@returns` tag.</span></span>

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

---
### <a name="streaming"></a><span data-ttu-id="b1013-205">@diffusionencontinu</span><span class="sxs-lookup"><span data-stu-id="b1013-205">@streaming</span></span>
<a id="streaming"/>

<span data-ttu-id="b1013-206">Utilisé pour indiquer qu’une fonction personnalisée est une fonction diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="b1013-206">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="b1013-207">Le dernier paramètre doit être de type `CustomFunctions.StreamingInvocation<ResultType>`.</span><span class="sxs-lookup"><span data-stu-id="b1013-207">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="b1013-208">La fonction doit renvoyer `void`.</span><span class="sxs-lookup"><span data-stu-id="b1013-208">The function should return `void`.</span></span>

<span data-ttu-id="b1013-209">Les fonctions de diffusion en continu ne renvoient pas de valeurs directement, mais doivent plutôt appeler `setResult(result: ResultType)` en utilisant le dernier paramètre.</span><span class="sxs-lookup"><span data-stu-id="b1013-209">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="b1013-210">Les exceptions levées par une fonction en continu sont ignorées.</span><span class="sxs-lookup"><span data-stu-id="b1013-210">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="b1013-211">`setResult()`peut être appelée avec Error pour indiquer un résultat erroné.</span><span class="sxs-lookup"><span data-stu-id="b1013-211">`setResult()` may be called with Error to indicate an error result.</span></span> <span data-ttu-id="b1013-212">Si vous souhaitez consulter un exemple de fonction de diffusion en continu et obtenir d’autres informations, veuillez vous reporter à la section [Créer une fonction de diffusion en continu](./custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="b1013-212">For an example of a streaming function and more information, see [Make a streaming function](./custom-functions-web-reqs.md#make-a-streaming-function).</span></span>

<span data-ttu-id="b1013-213">Les fonctions de diffusion en continu ne peuvent pas être marquées comme étant [@volatile](#volatile).</span><span class="sxs-lookup"><span data-stu-id="b1013-213">Streaming functions cannot be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a><span data-ttu-id="b1013-214">@volatile</span><span class="sxs-lookup"><span data-stu-id="b1013-214">@volatile</span></span>
<a id="volatile"/>

<span data-ttu-id="b1013-215">Une fonction volatile est une fonction dont le résultat peut changer d’un moment à l’autre, même si elle ne récupère pas d’argument ou si ses arguments ne changent pas.</span><span class="sxs-lookup"><span data-stu-id="b1013-215">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="b1013-216">À chaque calcul, Excel réévalue les cellules contenant des fonctions volatiles, ainsi que toutes leurs cellules dépendantes.</span><span class="sxs-lookup"><span data-stu-id="b1013-216">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="b1013-217">C’est pourquoi, un trop grand nombre de dépendances de fonctions volatiles risque de ralentir les calculs. Nous vous recommandons d’en utiliser aussi peu que possible.</span><span class="sxs-lookup"><span data-stu-id="b1013-217">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="b1013-218">Les fonctions de diffusion en continu ne peuvent pas être volatiles.</span><span class="sxs-lookup"><span data-stu-id="b1013-218">Streaming functions cannot be volatile.</span></span>

<span data-ttu-id="b1013-219">La fonction suivante est volatile et utilise la balise `@volatile`.</span><span class="sxs-lookup"><span data-stu-id="b1013-219">The following function is volatile and uses the `@volatile` tag.</span></span>

```js
/**
 * Simulates rolling a 6-sided dice.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

---

## <a name="types"></a><span data-ttu-id="b1013-220">Types</span><span class="sxs-lookup"><span data-stu-id="b1013-220">Types</span></span>

<span data-ttu-id="b1013-221">En spécifiant un type de paramètre, Excel convertit les valeurs en ce type, avant d’appeler la fonction.</span><span class="sxs-lookup"><span data-stu-id="b1013-221">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="b1013-222">Si le type est `any`, Excel n’effectue pas de conversion.</span><span class="sxs-lookup"><span data-stu-id="b1013-222">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="b1013-223">Types de valeur</span><span class="sxs-lookup"><span data-stu-id="b1013-223">Value types</span></span>

<span data-ttu-id="b1013-224">Une valeur unique peut être représentée à l’aide d’un des types suivants : `boolean`, `number` ou `string`.</span><span class="sxs-lookup"><span data-stu-id="b1013-224">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="b1013-225">Type matrice</span><span class="sxs-lookup"><span data-stu-id="b1013-225">Matrix type</span></span>

<span data-ttu-id="b1013-226">Utilisez une matrice à deux dimensions pour que le paramètre ou la valeur renvoyée soit une matrice de valeurs.</span><span class="sxs-lookup"><span data-stu-id="b1013-226">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="b1013-227">Par exemple, le type `number[][]` indique une matrice de nombres.</span><span class="sxs-lookup"><span data-stu-id="b1013-227">For example, the type `number[][]` indicates a matrix of numbers.</span></span> <span data-ttu-id="b1013-228">`string[][]`indique une matrice de chaînes.</span><span class="sxs-lookup"><span data-stu-id="b1013-228">`string[][]` indicates a matrix of strings.</span></span>

### <a name="error-type"></a><span data-ttu-id="b1013-229">Type d’erreur</span><span class="sxs-lookup"><span data-stu-id="b1013-229">Error type</span></span>

<span data-ttu-id="b1013-230">Une fonction qui n’est pas une fonction de diffusion en continu peut indiquer une erreur en renvoyant un type Error.</span><span class="sxs-lookup"><span data-stu-id="b1013-230">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="b1013-231">Une fonction de diffusion en continu peut indiquer une erreur en appelant`setResult()`avec un type Error.</span><span class="sxs-lookup"><span data-stu-id="b1013-231">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="b1013-232">Promise</span><span class="sxs-lookup"><span data-stu-id="b1013-232">Promise</span></span>

<span data-ttu-id="b1013-233">Une fonction peut renvoyer un objet Promise (pour « promesse »). Ce dernier fournit une valeur lors de sa résolution.</span><span class="sxs-lookup"><span data-stu-id="b1013-233">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="b1013-234">Si la résolution de l’objet Promise est refusée, cela entraîne une erreur.</span><span class="sxs-lookup"><span data-stu-id="b1013-234">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="b1013-235">Autres types</span><span class="sxs-lookup"><span data-stu-id="b1013-235">Other types</span></span>

<span data-ttu-id="b1013-236">Tout autre type sera traité comme une erreur.</span><span class="sxs-lookup"><span data-stu-id="b1013-236">Any other type will be treated as an error.</span></span>

## <a name="next-steps"></a><span data-ttu-id="b1013-237">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="b1013-237">Next steps</span></span>
<span data-ttu-id="b1013-238">Découvrez les [conventions d’affectation des noms des fonctions personnalisées](custom-functions-naming.md).</span><span class="sxs-lookup"><span data-stu-id="b1013-238">Learn about [naming conventions for custom functions](custom-functions-naming.md).</span></span> <span data-ttu-id="b1013-239">Découvrez également comment [localiser vos fonctions](custom-functions-localize.md), ce qui implique que vous [écriviez votre fichier JSON à la main](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="b1013-239">Alternatively, learn how to [localize your functions](custom-functions-localize.md) which requires you to [write your JSON file by hand](custom-functions-json.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="b1013-240">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b1013-240">See also</span></span>

* [<span data-ttu-id="b1013-241">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="b1013-241">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="b1013-242">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="b1013-242">Create custom functions in Excel</span></span>](custom-functions-overview.md)
