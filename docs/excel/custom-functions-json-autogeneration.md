---
ms.date: 12/22/2020
description: Utiliser les balises JSDOC pour créer dynamiquement vos fonctions personnalisées de métadonnées JSON.
title: Générer automatiquement des métadonnées JSON pour des fonctions personnalisées
localization_priority: Normal
ms.openlocfilehash: 4f7553646f51518fd4c5ed6c66081f910ec813a1
ms.sourcegitcommit: 48b9c3b63668b2a53ce73f92ce124ca07c5ca68c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/28/2020
ms.locfileid: "49735556"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a><span data-ttu-id="f28ad-103">Générer automatiquement des métadonnées JSON pour des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="f28ad-103">Autogenerate JSON metadata for custom functions</span></span>

<span data-ttu-id="f28ad-104">Si vous écrivez une fonction Excel personnalisée en JavaScript ou TypeScript, vous pouvez utiliser les [balises JSDoc](https://jsdoc.app/) pour la détailler en ajoutant des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="f28ad-104">When an Excel custom function is written in JavaScript or TypeScript, [JSDoc tags](https://jsdoc.app/) are used to provide extra information about the custom function.</span></span> <span data-ttu-id="f28ad-105">Les balises JSDoc sont ensuite utilisées lors de la génération pour créer le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="f28ad-105">The JSDoc tags are then used at build time to create the JSON metadata file.</span></span> <span data-ttu-id="f28ad-106">L’utilisation de balises JSDoc vous évite d’avoir à [modifier manuellement le fichier de métadonnées JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="f28ad-106">Using JSDoc tags saves you from the effort of [manually editing the JSON metadata file](custom-functions-json.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="f28ad-107">Ajoutez la balise `@customfunction` dans les commentaires du code d’une fonction JavaScript ou TypeScript pour indiquer qu’il s’agit d’une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="f28ad-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="f28ad-108">Vous pouvez fournir les types de paramètres de la fonction en utilisant la balise[@param](#param)dans JavaScript, ou en précisant le [type de fonction](https://www.typescriptlang.org/docs/handbook/functions.html) dans TypeScript.</span><span class="sxs-lookup"><span data-stu-id="f28ad-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="f28ad-109">Si vous voulez en savoir plus, veuillez consulter les sections relatives à la balise[@param](#param) et aux sections[types](#types).</span><span class="sxs-lookup"><span data-stu-id="f28ad-109">For more information, see the [@param](#param) tag and [Types](#types) sections.</span></span>

### <a name="adding-a-description-to-a-function"></a><span data-ttu-id="f28ad-110">Ajout d’une description à une fonction</span><span class="sxs-lookup"><span data-stu-id="f28ad-110">Adding a description to a function</span></span>

<span data-ttu-id="f28ad-111">La description s’affiche pour l’utilisateur sous forme de texte d’aide lorsqu’il a besoin d’aide pour comprendre le rôle de votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="f28ad-111">The description is displayed to the user as help text when they need help to understand what your custom function does.</span></span> <span data-ttu-id="f28ad-112">La description ne nécessite aucune balise spécifique.</span><span class="sxs-lookup"><span data-stu-id="f28ad-112">The description doesn't require any specific tag.</span></span> <span data-ttu-id="f28ad-113">Il vous suffit d’entrer une brève description dans le commentaire JSDoc.</span><span class="sxs-lookup"><span data-stu-id="f28ad-113">Just enter a short text description in the JSDoc comment.</span></span> <span data-ttu-id="f28ad-114">En général, la description est placée au début de la section commentaires JSDoc, mais elle fonctionnera peu importe son emplacement.</span><span class="sxs-lookup"><span data-stu-id="f28ad-114">In general the description is placed at the start of the JSDoc comment section, but it will work no matter where it is placed.</span></span>

<span data-ttu-id="f28ad-115">Pour consulter des exemples de descriptions de fonction intégrées, ouvrez Excel, accédez à l’onglet **formules** , puis sélectionnez **insérer une fonction**.</span><span class="sxs-lookup"><span data-stu-id="f28ad-115">To see examples of the built-in function descriptions, open Excel, go to the **Formulas** tab, and choose **Insert function**.</span></span> <span data-ttu-id="f28ad-116">Vous pouvez ensuite parcourir toutes les descriptions de fonction et voir vos propres fonctions personnalisées répertoriées.</span><span class="sxs-lookup"><span data-stu-id="f28ad-116">You can then browse through all the function descriptions, and also see your own custom functions listed.</span></span>

<span data-ttu-id="f28ad-117">Dans cet exemple, la phrase «calcule le volume d’une sphère.»</span><span class="sxs-lookup"><span data-stu-id="f28ad-117">In the following example, the phrase "Calculates the volume of a sphere."</span></span> <span data-ttu-id="f28ad-118">est la description de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="f28ad-118">is the description for the custom function.</span></span>

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```


## <a name="jsdoc-tags"></a><span data-ttu-id="f28ad-119">Balises JSDoc</span><span class="sxs-lookup"><span data-stu-id="f28ad-119">JSDoc Tags</span></span>

<span data-ttu-id="f28ad-120">Les balises JSDoc suivantes sont prises en charge dans les fonctions personnalisées Excel.</span><span class="sxs-lookup"><span data-stu-id="f28ad-120">The following JSDoc tags are supported in Excel custom functions.</span></span>

* [<span data-ttu-id="f28ad-121">@ annulable</span><span class="sxs-lookup"><span data-stu-id="f28ad-121">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="f28ad-122">[@fonctionpersonnalisée](#customfunction)nom id</span><span class="sxs-lookup"><span data-stu-id="f28ad-122">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="f28ad-123">url[@urlaide](#helpurl)</span><span class="sxs-lookup"><span data-stu-id="f28ad-123">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="f28ad-124">[@param](#param) _{type}_ description nom</span><span class="sxs-lookup"><span data-stu-id="f28ad-124">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="f28ad-125">@requièreuneadresse</span><span class="sxs-lookup"><span data-stu-id="f28ad-125">@requiresAddress</span></span>](#requiresAddress)
* <span data-ttu-id="f28ad-126">[@renvoie](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="f28ad-126">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="f28ad-127">@diffusionencontinu</span><span class="sxs-lookup"><span data-stu-id="f28ad-127">@streaming</span></span>](#streaming)
* [<span data-ttu-id="f28ad-128">@volatile</span><span class="sxs-lookup"><span data-stu-id="f28ad-128">@volatile</span></span>](#volatile)

---
<a id="cancelable"></a>

### <a name="cancelable"></a><span data-ttu-id="f28ad-129">@ annulable</span><span class="sxs-lookup"><span data-stu-id="f28ad-129">@cancelable</span></span>

<span data-ttu-id="f28ad-130">Indique qu’une fonction personnalisée effectue une action lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="f28ad-130">Indicates that a custom function performs an action when the function is canceled.</span></span>

<span data-ttu-id="f28ad-131">Le dernier paramètre de la fonction doit être de type `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="f28ad-131">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="f28ad-132">La fonction peut affecter une fonction à la `oncanceled` propriété pour indiquer le résultat lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="f28ad-132">The function can assign a function to the `oncanceled` property to denote the result when the function is canceled.</span></span>

<span data-ttu-id="f28ad-133">Si le dernier paramètre de fonction est de type `CustomFunctions.CancelableInvocation`, il sera considéré comme `@cancelable`, même si la balise n’apparaît pas.</span><span class="sxs-lookup"><span data-stu-id="f28ad-133">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag isn't present.</span></span>

<span data-ttu-id="f28ad-134">Une fonction ne peut pas contenir les deux balises `@cancelable` et `@streaming`.</span><span class="sxs-lookup"><span data-stu-id="f28ad-134">A function can't have both `@cancelable` and `@streaming` tags.</span></span>

<a id="customfunction"></a>

### <a name="customfunction"></a><span data-ttu-id="f28ad-135">@fonctionpersonnalisée</span><span class="sxs-lookup"><span data-stu-id="f28ad-135">@customfunction</span></span>

<span data-ttu-id="f28ad-136">Syntaxe: @fonctionpersonnalisée _id_ _nom_</span><span class="sxs-lookup"><span data-stu-id="f28ad-136">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="f28ad-137">Cette balise indique que la fonction JavaScript/dactylographié est une fonction personnalisée Excel.</span><span class="sxs-lookup"><span data-stu-id="f28ad-137">This tag indicates that the JavaScript/TypeScript function is an Excel custom function.</span></span> <span data-ttu-id="f28ad-138">Il est nécessaire de créer des métadonnées pour la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="f28ad-138">It is required to create metadata for the custom function.</span></span>

<span data-ttu-id="f28ad-139">Voici un exemple de cette balise.</span><span class="sxs-lookup"><span data-stu-id="f28ad-139">The following shows an example of this tag.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a><span data-ttu-id="f28ad-140">id</span><span class="sxs-lookup"><span data-stu-id="f28ad-140">id</span></span>

<span data-ttu-id="f28ad-141">L' `id` identifie une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="f28ad-141">The `id` identifies a custom function.</span></span>

* <span data-ttu-id="f28ad-142">Si`id`n’est pas fourni, le nom de la fonction JavaScript/TypeScript est converti en majuscules, et les caractères rejetés sont supprimés.</span><span class="sxs-lookup"><span data-stu-id="f28ad-142">If `id` isn't provided, the JavaScript/TypeScript function name is converted to uppercase and disallowed characters are removed.</span></span>
* <span data-ttu-id="f28ad-143">Le `id`doit être unique pour toutes les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="f28ad-143">The `id` must be unique for all custom functions.</span></span>
* <span data-ttu-id="f28ad-144">Les caractères autorisés sont les suivants : A-Z, a-z, 0-9, traits de soulignement (\_) et point (.).</span><span class="sxs-lookup"><span data-stu-id="f28ad-144">The allowed characters are limited to: A-Z, a-z, 0-9, underscores (\_), and period (.).</span></span>

<span data-ttu-id="f28ad-145">Dans l’exemple suivant, Increments correspond à l’`id` et au `name` de la fonction.</span><span class="sxs-lookup"><span data-stu-id="f28ad-145">In the following example, increment is the `id` and the `name` of the function.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INCREMENT
 * ...
 */
```

#### <a name="name"></a><span data-ttu-id="f28ad-146">name</span><span class="sxs-lookup"><span data-stu-id="f28ad-146">name</span></span>

<span data-ttu-id="f28ad-147">Fournit le nom d’affichage `name`de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="f28ad-147">Provides the display `name` for the custom function.</span></span>

* <span data-ttu-id="f28ad-148">Si aucun nom n’est fourni, l’id servira aussi de nom.</span><span class="sxs-lookup"><span data-stu-id="f28ad-148">If name isn't provided, the id is also used as the name.</span></span>
* <span data-ttu-id="f28ad-149">Caractères autorisés : [caractères alphanumériques Unicode](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic) (lettres, chiffres), point (.) et trait de soulignement (\_).</span><span class="sxs-lookup"><span data-stu-id="f28ad-149">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="f28ad-150">Doit commencer par une lettre.</span><span class="sxs-lookup"><span data-stu-id="f28ad-150">Must start with a letter.</span></span>
* <span data-ttu-id="f28ad-151">Sa longueur maximale est limitée à 128 caractères.</span><span class="sxs-lookup"><span data-stu-id="f28ad-151">Maximum length is 128 characters.</span></span>

<span data-ttu-id="f28ad-152">Dans l’exemple suivant, INC correspond à l’`id` de la fonction, tandis que `increment` correspond au `name`.</span><span class="sxs-lookup"><span data-stu-id="f28ad-152">In the following example, INC is the `id` of the function and `increment` is the `name`.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a><span data-ttu-id="f28ad-153">description</span><span class="sxs-lookup"><span data-stu-id="f28ad-153">description</span></span>

<span data-ttu-id="f28ad-154">Une description apparaît pour les utilisateurs dans Excel lorsqu’ils entrent dans la fonction et spécifie le rôle de la fonction.</span><span class="sxs-lookup"><span data-stu-id="f28ad-154">A description appears to users in Excel as they are entering the function and specifies what the function does.</span></span> <span data-ttu-id="f28ad-155">Une description ne nécessite aucune balise spécifique.</span><span class="sxs-lookup"><span data-stu-id="f28ad-155">A description doesn't require any specific tag.</span></span> <span data-ttu-id="f28ad-156">Ajoutez une description à une fonction personnalisée en ajoutant une expression pour décrire le rôle de la fonction dans le commentaire JSDoc.</span><span class="sxs-lookup"><span data-stu-id="f28ad-156">Add a description to a custom function by adding a phrase to describe what the function does inside the JSDoc comment.</span></span> <span data-ttu-id="f28ad-157">Par défaut, le texte non balisé dans la section commentaire JSDoc est la description de la fonction.</span><span class="sxs-lookup"><span data-stu-id="f28ad-157">By default, whatever text is untagged in the JSDoc comment section will be the description of the function.</span></span>

<span data-ttu-id="f28ad-158">Dans l’exemple suivant, la phrase « A function that adds two numbers » (« Une fonction qui ajoute deux nombres ») est la description de la fonction personnalisée dont la propriété ID est `ADD`.</span><span class="sxs-lookup"><span data-stu-id="f28ad-158">In the following example, the phrase "A function that adds two numbers" is the description for the custom function with the id property of `ADD`.</span></span>

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

<a id="helpurl"></a>

### <a name="helpurl"></a><span data-ttu-id="f28ad-159">@urlaide</span><span class="sxs-lookup"><span data-stu-id="f28ad-159">@helpurl</span></span>

<span data-ttu-id="f28ad-160">Syntaxe: @urlaide _url_</span><span class="sxs-lookup"><span data-stu-id="f28ad-160">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="f28ad-161">L’_url_ fournie est affichée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="f28ad-161">The provided _url_ is displayed in Excel.</span></span>

<span data-ttu-id="f28ad-162">Dans l’exemple suivant, le `helpurl` est `www.contoso.com/weatherhelp` .</span><span class="sxs-lookup"><span data-stu-id="f28ad-162">In the following example, the `helpurl` is `www.contoso.com/weatherhelp`.</span></span>

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

<a id="param"></a>

### <a name="param"></a><span data-ttu-id="f28ad-163">@param</span><span class="sxs-lookup"><span data-stu-id="f28ad-163">@param</span></span>

#### <a name="javascript"></a><span data-ttu-id="f28ad-164">JavaScript</span><span class="sxs-lookup"><span data-stu-id="f28ad-164">JavaScript</span></span>

<span data-ttu-id="f28ad-165">Syntaxe JavaScript : @param {type} nom _description_</span><span class="sxs-lookup"><span data-stu-id="f28ad-165">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="f28ad-166">`{type}` spécifie les informations de type entre accolades.</span><span class="sxs-lookup"><span data-stu-id="f28ad-166">`{type}` specifies the type info within curly braces.</span></span> <span data-ttu-id="f28ad-167">Consultez la section [Types](#types) pour savoir quels types peuvent être utilisés.</span><span class="sxs-lookup"><span data-stu-id="f28ad-167">See the [Types](#types) section for more information about the types which may be used.</span></span> <span data-ttu-id="f28ad-168">Si aucun type n’est spécifié, le type par défaut est `any` utilisé.</span><span class="sxs-lookup"><span data-stu-id="f28ad-168">If no type is specified, the default type `any` will be used.</span></span>
* <span data-ttu-id="f28ad-169">`name` Spécifie le paramètre auquel s’applique la balise @param.</span><span class="sxs-lookup"><span data-stu-id="f28ad-169">`name` specifies the parameter that the @param tag applies to.</span></span> <span data-ttu-id="f28ad-170">Elle est obligatoire.</span><span class="sxs-lookup"><span data-stu-id="f28ad-170">It is required.</span></span>
* <span data-ttu-id="f28ad-171">`description`fournit la description qui s’affiche dans Excel pour le paramètre de la fonction.</span><span class="sxs-lookup"><span data-stu-id="f28ad-171">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="f28ad-172">Elle est facultative.</span><span class="sxs-lookup"><span data-stu-id="f28ad-172">It is optional.</span></span>

<span data-ttu-id="f28ad-173">Pour désigner un paramètre de fonction personnalisée comme étant facultatif :</span><span class="sxs-lookup"><span data-stu-id="f28ad-173">To denote a custom function parameter as optional:</span></span>

* <span data-ttu-id="f28ad-174">Placez les crochets autour du nom du paramètre.</span><span class="sxs-lookup"><span data-stu-id="f28ad-174">Put square brackets around the parameter name.</span></span> <span data-ttu-id="f28ad-175">Par exemple : `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="f28ad-175">For example: `@param {string} [text] Optional text`.</span></span>

> [!NOTE]
> <span data-ttu-id="f28ad-176">La valeur par défaut pour les paramètres facultatifs est `null`.</span><span class="sxs-lookup"><span data-stu-id="f28ad-176">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="f28ad-177">L’exemple suivant montre une fonction ADD qui ajoute deux ou trois nombres, le troisième étant un paramètre facultatif.</span><span class="sxs-lookup"><span data-stu-id="f28ad-177">The following example shows a ADD function that adds two or three numbers, with the third number as an optional parameter.</span></span>

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

#### <a name="typescript"></a><span data-ttu-id="f28ad-178">TypeScript</span><span class="sxs-lookup"><span data-stu-id="f28ad-178">TypeScript</span></span>

<span data-ttu-id="f28ad-179">Syntaxe TypeScript : nom @param _description_</span><span class="sxs-lookup"><span data-stu-id="f28ad-179">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="f28ad-180">`name` Spécifie le paramètre auquel s’applique la balise @param.</span><span class="sxs-lookup"><span data-stu-id="f28ad-180">`name` specifies the parameter that the @param tag applies to.</span></span> <span data-ttu-id="f28ad-181">Elle est obligatoire.</span><span class="sxs-lookup"><span data-stu-id="f28ad-181">It is required.</span></span>
* <span data-ttu-id="f28ad-182">`description`fournit la description qui s’affiche dans Excel pour le paramètre de la fonction.</span><span class="sxs-lookup"><span data-stu-id="f28ad-182">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="f28ad-183">Elle est facultative.</span><span class="sxs-lookup"><span data-stu-id="f28ad-183">It is optional.</span></span>

<span data-ttu-id="f28ad-184">Consultez la section [Types](#types) pour savoir quels types de paramètres de fonction peuvent être utilisés.</span><span class="sxs-lookup"><span data-stu-id="f28ad-184">See the [Types](#types) section for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="f28ad-185">Pour désigner un paramètre de fonction personnalisée comme étant facultatif, effectuez l’une des actions suivantes :</span><span class="sxs-lookup"><span data-stu-id="f28ad-185">To denote a custom function parameter as optional, do one of the following:</span></span>

* <span data-ttu-id="f28ad-186">Utilisez un paramètre facultatif.</span><span class="sxs-lookup"><span data-stu-id="f28ad-186">Use an optional parameter.</span></span> <span data-ttu-id="f28ad-187">Par exemple : `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="f28ad-187">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="f28ad-188">Définissez ce paramètre sur une valeur par défaut.</span><span class="sxs-lookup"><span data-stu-id="f28ad-188">Give the parameter a default value.</span></span> <span data-ttu-id="f28ad-189">Par exemple : `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="f28ad-189">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="f28ad-190">Pour consulter une description détaillée du @param, reportez-vous à la page suivante : [JSDoc](https://jsdoc.app/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="f28ad-190">For detailed description of the @param see: [JSDoc](https://jsdoc.app/tags-param.html)</span></span>

> [!NOTE]
> <span data-ttu-id="f28ad-191">La valeur par défaut pour les paramètres facultatifs est `null`.</span><span class="sxs-lookup"><span data-stu-id="f28ad-191">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="f28ad-192">L’exemple suivant représente la fonction `add` qui ajoute deux nombres.</span><span class="sxs-lookup"><span data-stu-id="f28ad-192">The following example shows the `add` function that adds two numbers.</span></span>

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

<a id="requiresAddress"></a>

### <a name="requiresaddress"></a><span data-ttu-id="f28ad-193">@requièreuneadresse</span><span class="sxs-lookup"><span data-stu-id="f28ad-193">@requiresAddress</span></span>

<span data-ttu-id="f28ad-194">Indique que l’adresse de la cellule dans laquelle la fonction est évaluée doit être fournie.</span><span class="sxs-lookup"><span data-stu-id="f28ad-194">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span>

<span data-ttu-id="f28ad-195">Le dernier paramètre de la fonction doit être de type `CustomFunctions.Invocation` ou un type dérivé.</span><span class="sxs-lookup"><span data-stu-id="f28ad-195">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="f28ad-196">Lorsque la fonction est appelée, la propriété `address` contiendra l’adresse.</span><span class="sxs-lookup"><span data-stu-id="f28ad-196">When the function is called, the `address` property will contain the address.</span></span>

<span data-ttu-id="f28ad-197">L’exemple suivant montre comment utiliser le `invocation` paramètre en combinaison avec `@requiresAddress` pour renvoyer l’adresse de la cellule qui a appelé votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="f28ad-197">The following sample shows how to use the `invocation` parameter in combination with `@requiresAddress` to return the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="f28ad-198">Pour plus d’informations, consultez la rubrique [paramètre invocation](custom-functions-parameter-options.md#invocation-parameter) .</span><span class="sxs-lookup"><span data-stu-id="f28ad-198">See [Invocation parameter](custom-functions-parameter-options.md#invocation-parameter) for more information.</span></span>

```js
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
function getAddress(first, second, invocation) {
  var address = invocation.address;
  return address;
}
```

<a id="returns"></a>

### <a name="returns"></a><span data-ttu-id="f28ad-199">@renvoie :</span><span class="sxs-lookup"><span data-stu-id="f28ad-199">@returns</span></span>

<span data-ttu-id="f28ad-200">Syntaxe: @renvoie {_type_}</span><span class="sxs-lookup"><span data-stu-id="f28ad-200">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="f28ad-201">Fournit le type pour la valeur renvoyée.</span><span class="sxs-lookup"><span data-stu-id="f28ad-201">Provides the type for the return value.</span></span>

<span data-ttu-id="f28ad-202">Si `{type}` est omis, les informations de type TypeScript seront utilisées.</span><span class="sxs-lookup"><span data-stu-id="f28ad-202">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="f28ad-203">S’il n’existe aucune information définissant le type, ce dernier sera `any`.</span><span class="sxs-lookup"><span data-stu-id="f28ad-203">If there is no type info, the type will be `any`.</span></span>

<span data-ttu-id="f28ad-204">L’exemple suivant représente la fonction `add` qui utilise la balise `@returns`.</span><span class="sxs-lookup"><span data-stu-id="f28ad-204">The following example shows the `add` function that uses the `@returns` tag.</span></span>

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

<a id="streaming"></a>

### <a name="streaming"></a><span data-ttu-id="f28ad-205">@diffusionencontinu</span><span class="sxs-lookup"><span data-stu-id="f28ad-205">@streaming</span></span>

<span data-ttu-id="f28ad-206">Utilisé pour indiquer qu’une fonction personnalisée est une fonction diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="f28ad-206">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="f28ad-207">Le dernier paramètre est de type `CustomFunctions.StreamingInvocation<ResultType>` .</span><span class="sxs-lookup"><span data-stu-id="f28ad-207">The last parameter is of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="f28ad-208">La fonction renvoie `void` .</span><span class="sxs-lookup"><span data-stu-id="f28ad-208">The function returns `void`.</span></span>

<span data-ttu-id="f28ad-209">Les fonctions de diffusion en continu ne renvoient pas directement de valeurs, mais appelent `setResult(result: ResultType)` à l’aide du dernier paramètre.</span><span class="sxs-lookup"><span data-stu-id="f28ad-209">Streaming functions don't return values directly, instead they call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="f28ad-210">Les exceptions levées par une fonction en continu sont ignorées.</span><span class="sxs-lookup"><span data-stu-id="f28ad-210">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="f28ad-211">`setResult()`peut être appelée avec Error pour indiquer un résultat erroné.</span><span class="sxs-lookup"><span data-stu-id="f28ad-211">`setResult()` may be called with Error to indicate an error result.</span></span> <span data-ttu-id="f28ad-212">Si vous souhaitez consulter un exemple de fonction de diffusion en continu et obtenir d’autres informations, veuillez vous reporter à la section [Créer une fonction de diffusion en continu](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="f28ad-212">For an example of a streaming function and more information, see [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span>

<span data-ttu-id="f28ad-213">Les fonctions de diffusion en continu ne peuvent pas être marquées comme étant [@volatile](#volatile).</span><span class="sxs-lookup"><span data-stu-id="f28ad-213">Streaming functions can't be marked as [@volatile](#volatile).</span></span>

<a id="volatile"></a>

### <a name="volatile"></a><span data-ttu-id="f28ad-214">@volatile</span><span class="sxs-lookup"><span data-stu-id="f28ad-214">@volatile</span></span>

<span data-ttu-id="f28ad-215">Une fonction volatile est une fonction dont le résultat peut changer d’un moment à l’autre, même si elle ne récupère pas d’argument ou si ses arguments ne changent pas.</span><span class="sxs-lookup"><span data-stu-id="f28ad-215">A volatile function is one whose result isn't the same from one moment to the next, even if it takes no arguments or the arguments haven't changed.</span></span> <span data-ttu-id="f28ad-216">À chaque calcul, Excel réévalue les cellules contenant des fonctions volatiles, ainsi que toutes leurs cellules dépendantes.</span><span class="sxs-lookup"><span data-stu-id="f28ad-216">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="f28ad-217">C’est pourquoi, un trop grand nombre de dépendances de fonctions volatiles risque de ralentir les calculs. Nous vous recommandons d’en utiliser aussi peu que possible.</span><span class="sxs-lookup"><span data-stu-id="f28ad-217">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="f28ad-218">Les fonctions de diffusion en continu ne peuvent pas être volatiles.</span><span class="sxs-lookup"><span data-stu-id="f28ad-218">Streaming functions can't be volatile.</span></span>

<span data-ttu-id="f28ad-219">La fonction suivante est volatile et utilise la balise `@volatile`.</span><span class="sxs-lookup"><span data-stu-id="f28ad-219">The following function is volatile and uses the `@volatile` tag.</span></span>

```js
/**
 * Simulates rolling a 6-sided die.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

---

## <a name="types"></a><span data-ttu-id="f28ad-220">Types</span><span class="sxs-lookup"><span data-stu-id="f28ad-220">Types</span></span>

<span data-ttu-id="f28ad-221">En spécifiant un type de paramètre, Excel convertit les valeurs en ce type, avant d’appeler la fonction.</span><span class="sxs-lookup"><span data-stu-id="f28ad-221">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="f28ad-222">Si le type est `any`, Excel n’effectue pas de conversion.</span><span class="sxs-lookup"><span data-stu-id="f28ad-222">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="f28ad-223">Types de valeur</span><span class="sxs-lookup"><span data-stu-id="f28ad-223">Value types</span></span>

<span data-ttu-id="f28ad-224">Une valeur unique peut être représentée à l’aide d’un des types suivants : `boolean`, `number` ou `string`.</span><span class="sxs-lookup"><span data-stu-id="f28ad-224">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="f28ad-225">Type matrice</span><span class="sxs-lookup"><span data-stu-id="f28ad-225">Matrix type</span></span>

<span data-ttu-id="f28ad-226">Utilisez une matrice à deux dimensions pour que le paramètre ou la valeur renvoyée soit une matrice de valeurs.</span><span class="sxs-lookup"><span data-stu-id="f28ad-226">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="f28ad-227">Par exemple, le type `number[][]` indique une matrice de nombres.</span><span class="sxs-lookup"><span data-stu-id="f28ad-227">For example, the type `number[][]` indicates a matrix of numbers.</span></span> <span data-ttu-id="f28ad-228">`string[][]`indique une matrice de chaînes.</span><span class="sxs-lookup"><span data-stu-id="f28ad-228">`string[][]` indicates a matrix of strings.</span></span>

### <a name="error-type"></a><span data-ttu-id="f28ad-229">Type d’erreur</span><span class="sxs-lookup"><span data-stu-id="f28ad-229">Error type</span></span>

<span data-ttu-id="f28ad-230">Une fonction qui n’est pas une fonction de diffusion en continu peut indiquer une erreur en renvoyant un type Error.</span><span class="sxs-lookup"><span data-stu-id="f28ad-230">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="f28ad-231">Une fonction de diffusion en continu peut indiquer une erreur en appelant`setResult()`avec un type Error.</span><span class="sxs-lookup"><span data-stu-id="f28ad-231">A streaming function can indicate an error by calling `setResult()` with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="f28ad-232">Promise</span><span class="sxs-lookup"><span data-stu-id="f28ad-232">Promise</span></span>

<span data-ttu-id="f28ad-233">Une fonction peut renvoyer une promesse, qui fournit la valeur lorsque la promesse est résolue.</span><span class="sxs-lookup"><span data-stu-id="f28ad-233">A function can return a Promise, that provides the value when the promise is resolved.</span></span> <span data-ttu-id="f28ad-234">Si la promesse est rejetée, elle génère une erreur.</span><span class="sxs-lookup"><span data-stu-id="f28ad-234">If the promise is rejected, then it will throw an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="f28ad-235">Autres types</span><span class="sxs-lookup"><span data-stu-id="f28ad-235">Other types</span></span>

<span data-ttu-id="f28ad-236">Tout autre type sera traité comme une erreur.</span><span class="sxs-lookup"><span data-stu-id="f28ad-236">Any other type will be treated as an error.</span></span>

## <a name="next-steps"></a><span data-ttu-id="f28ad-237">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="f28ad-237">Next steps</span></span>

<span data-ttu-id="f28ad-238">Découvrez les [conventions d’affectation des noms des fonctions personnalisées](custom-functions-naming.md).</span><span class="sxs-lookup"><span data-stu-id="f28ad-238">Learn about [naming conventions for custom functions](custom-functions-naming.md).</span></span> <span data-ttu-id="f28ad-239">Découvrez également comment [localiser vos fonctions](custom-functions-localize.md), ce qui implique que vous [écriviez votre fichier JSON à la main](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="f28ad-239">Alternatively, learn how to [localize your functions](custom-functions-localize.md) which requires you to [write your JSON file by hand](custom-functions-json.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="f28ad-240">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f28ad-240">See also</span></span>

* [<span data-ttu-id="f28ad-241">Créer manuellement des métadonnées JSON pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="f28ad-241">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="f28ad-242">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="f28ad-242">Create custom functions in Excel</span></span>](custom-functions-overview.md)
