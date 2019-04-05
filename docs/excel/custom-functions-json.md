---
ms.date: 03/29/2019
description: Définissez des métadonnées pour des fonctions personnalisées dans Excel.
title: Métadonnées pour des fonctions personnalisées dans Excel (aperçu)
localization_priority: Normal
ms.openlocfilehash: 28a9a0207f7439af164eb9ca7c4b9ed9e966b3ed
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/04/2019
ms.locfileid: "31477550"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="a8837-103">Métadonnées de fonctions personnalisées (aperçu)</span><span class="sxs-lookup"><span data-stu-id="a8837-103">Custom functions metadata (preview)</span></span>

<span data-ttu-id="a8837-104">Lorsque vous définissez des [fonctions personnalisées](custom-functions-overview.md) dans votre complément Excel, votre projet de complément inclut un fichier de métadonnées JSON qui fournit les informations dont Excel a besoin pour enregistrer les fonctions personnalisées et les rendre accessibles aux utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="a8837-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project includes a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="a8837-105">Ce fichier est généré:</span><span class="sxs-lookup"><span data-stu-id="a8837-105">This file is generated either:</span></span>

- <span data-ttu-id="a8837-106">par vous-même, dans un fichier JSON manuscrit</span><span class="sxs-lookup"><span data-stu-id="a8837-106">by you, in a handwritten JSON file</span></span>
- <span data-ttu-id="a8837-107">à partir des commentaires JSDoc que vous entrez au début de votre fonction</span><span class="sxs-lookup"><span data-stu-id="a8837-107">from the JSDoc comments you enter at the beginning of your function</span></span>

<span data-ttu-id="a8837-108">Les fonctions personnalisées sont inscrites lorsque l'utilisateur exécute le complément pour la première fois et après qu'il est disponible pour le même utilisateur dans tous les classeurs.</span><span class="sxs-lookup"><span data-stu-id="a8837-108">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

<span data-ttu-id="a8837-109">Cet article décrit le format du fichier de métadonnées JSON, en supposant que vous l'écrivez manuellement.</span><span class="sxs-lookup"><span data-stu-id="a8837-109">This article describes the format of the JSON metadata file, assuming you are writing it by hand.</span></span> <span data-ttu-id="a8837-110">Pour plus d'informations sur la génération de fichiers JSON de commentaire JSDoc, voir [generate JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="a8837-110">For information about JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="a8837-111">Pour plus d’informations sur les autres fichiers à inclure dans votre projet de complément afin d’activer les fonctions personnalisées, voir [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="a8837-111">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

> <span data-ttu-id="a8837-112">Les paramètres du serveur qui héberge le fichier JSON doivent avoir [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) activée afin que les fonctions personnalisées s’exécutent correctement dans Excel Online.</span><span class="sxs-lookup"><span data-stu-id="a8837-112">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

## <a name="example-metadata"></a><span data-ttu-id="a8837-113">Exemple de métadonnées</span><span class="sxs-lookup"><span data-stu-id="a8837-113">Example metadata</span></span>

<span data-ttu-id="a8837-114">L’exemple suivant montre le contenu d’un fichier de métadonnées JSON pour un complément qui définit des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="a8837-114">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="a8837-115">Les sections qui suivent cet exemple fournissent des informations détaillées sur les propriétés individuelles au sein de cet exemple JSON.</span><span class="sxs-lookup"><span data-stu-id="a8837-115">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

```json
{
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      "description": "Add two numbers",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "number",
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "first",
          "description": "first number to add",
          "type": "number",
          "dimensionality": "scalar"
        },
        {
          "name": "second",
          "description": "second number to add",
          "type": "number",
          "dimensionality": "scalar"
        }
      ]
    },
    {
      "id": "GETDAY",
      "name": "GETDAY",
      "description": "Get the day of the week",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "string"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE", 
      "description":  "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "number",
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "increment",
          "description": "the number to be added each time",
          "type": "number",
          "dimensionality": "scalar"
        }
      ],
      "options": {
        "stream": true,
        "cancelable": true
      }
    },
    {
      "id": "SECONDHIGHEST",
      "name": "SECONDHIGHEST", 
      "description":  "Get the second highest number from a range",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "number",
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "range",
          "description": "the input range",
          "type": "number",
          "dimensionality": "matrix"
        }
      ]
    }
  ]
}
```

> [!NOTE]
> <span data-ttu-id="a8837-116">Un exemple de fichier JSON complet est disponible dans le dépôt GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json).</span><span class="sxs-lookup"><span data-stu-id="a8837-116">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="a8837-117">fonctions</span><span class="sxs-lookup"><span data-stu-id="a8837-117">functions</span></span> 

<span data-ttu-id="a8837-118">La propriété `functions` est un tableau d’objets de fonction personnalisés.</span><span class="sxs-lookup"><span data-stu-id="a8837-118">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="a8837-119">Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="a8837-119">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="a8837-120">Propriété</span><span class="sxs-lookup"><span data-stu-id="a8837-120">Property</span></span>  |  <span data-ttu-id="a8837-121">Type de données</span><span class="sxs-lookup"><span data-stu-id="a8837-121">Data type</span></span>  |  <span data-ttu-id="a8837-122">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="a8837-122">Required</span></span>  |  <span data-ttu-id="a8837-123">Description</span><span class="sxs-lookup"><span data-stu-id="a8837-123">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="a8837-124">string</span><span class="sxs-lookup"><span data-stu-id="a8837-124">string</span></span>  |  <span data-ttu-id="a8837-125">Non</span><span class="sxs-lookup"><span data-stu-id="a8837-125">No</span></span>  |  <span data-ttu-id="a8837-126">Description de la fonction que voient les utilisateurs finaux dans Excel.</span><span class="sxs-lookup"><span data-stu-id="a8837-126">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="a8837-127">Par exemple, **convertit une valeur Celsius en valeur Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="a8837-127">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="a8837-128">string</span><span class="sxs-lookup"><span data-stu-id="a8837-128">string</span></span>  |   <span data-ttu-id="a8837-129">Non</span><span class="sxs-lookup"><span data-stu-id="a8837-129">No</span></span>  |  <span data-ttu-id="a8837-130">URL fournissant des informations sur la fonction</span><span class="sxs-lookup"><span data-stu-id="a8837-130">URL that provides information about the function.</span></span> <span data-ttu-id="a8837-131">(elle est affichée dans un volet des tâches). Par exemple, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="a8837-131">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="a8837-132">string</span><span class="sxs-lookup"><span data-stu-id="a8837-132">string</span></span> | <span data-ttu-id="a8837-133">Oui</span><span class="sxs-lookup"><span data-stu-id="a8837-133">Yes</span></span> | <span data-ttu-id="a8837-134">Un ID unique pour la fonction.</span><span class="sxs-lookup"><span data-stu-id="a8837-134">A unique ID for the function.</span></span> <span data-ttu-id="a8837-135">Cet ID peut contenir uniquement des points et caractères alphanumériques et ne doit pas être modifié une fois défini.</span><span class="sxs-lookup"><span data-stu-id="a8837-135">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="a8837-136">string</span><span class="sxs-lookup"><span data-stu-id="a8837-136">string</span></span>  |  <span data-ttu-id="a8837-137">Oui</span><span class="sxs-lookup"><span data-stu-id="a8837-137">Yes</span></span>  |  <span data-ttu-id="a8837-138">Nom de la fonction que voient les utilisateurs finaux dans Excel.</span><span class="sxs-lookup"><span data-stu-id="a8837-138">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="a8837-139">Dans Excel, le nom de la fonction sera précédé de l’espace de noms de fonctions personnalisées spécifié dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="a8837-139">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="a8837-140">object</span><span class="sxs-lookup"><span data-stu-id="a8837-140">object</span></span>  |  <span data-ttu-id="a8837-141">Non</span><span class="sxs-lookup"><span data-stu-id="a8837-141">No</span></span>  |  <span data-ttu-id="a8837-142">Vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction.</span><span class="sxs-lookup"><span data-stu-id="a8837-142">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="a8837-143">Reportez-vous aux [options](#options) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="a8837-143">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="a8837-144">tableau</span><span class="sxs-lookup"><span data-stu-id="a8837-144">array</span></span>  |  <span data-ttu-id="a8837-145">Oui</span><span class="sxs-lookup"><span data-stu-id="a8837-145">Yes</span></span>  |  <span data-ttu-id="a8837-146">Tableau qui définit les paramètres d’entrée de la fonction.</span><span class="sxs-lookup"><span data-stu-id="a8837-146">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="a8837-147">Reportez-vous aux [paramètres](#parameters) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="a8837-147">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="a8837-148">objet</span><span class="sxs-lookup"><span data-stu-id="a8837-148">object</span></span>  |  <span data-ttu-id="a8837-149">Oui</span><span class="sxs-lookup"><span data-stu-id="a8837-149">Yes</span></span>  |  <span data-ttu-id="a8837-150">Objet qui définit le type d’informations renvoyées par la fonction.</span><span class="sxs-lookup"><span data-stu-id="a8837-150">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="a8837-151">Reportez-vous au [résultat](#result) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="a8837-151">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="a8837-152">options</span><span class="sxs-lookup"><span data-stu-id="a8837-152">options</span></span>

<span data-ttu-id="a8837-153">L’objet `options` vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction.</span><span class="sxs-lookup"><span data-stu-id="a8837-153">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="a8837-154">Le tableau suivant répertorie les propriétés de l’objet `options`.</span><span class="sxs-lookup"><span data-stu-id="a8837-154">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="a8837-155">Propriété</span><span class="sxs-lookup"><span data-stu-id="a8837-155">Property</span></span>  |  <span data-ttu-id="a8837-156">Type de données</span><span class="sxs-lookup"><span data-stu-id="a8837-156">Data type</span></span>  |  <span data-ttu-id="a8837-157">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="a8837-157">Required</span></span>  |  <span data-ttu-id="a8837-158">Description</span><span class="sxs-lookup"><span data-stu-id="a8837-158">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="a8837-159">boolean</span><span class="sxs-lookup"><span data-stu-id="a8837-159">boolean</span></span>  |  <span data-ttu-id="a8837-160">Non</span><span class="sxs-lookup"><span data-stu-id="a8837-160">No</span></span><br/><br/><span data-ttu-id="a8837-161">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="a8837-161">Default value is `false`.</span></span>  |  <span data-ttu-id="a8837-162">Si la valeur est `true`, Excel appelle le gestionnaire `onCanceled` chaque fois que l’utilisateur effectue une action ayant pour effet d’annuler la fonction, par exemple, en déclenchant manuellement un recalcul ou en modifiant une cellule référencée par la fonction.</span><span class="sxs-lookup"><span data-stu-id="a8837-162">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="a8837-163">Si vous utilisez cette option, Excel appelle la fonction JavaScript avec un paramètre `caller` supplémentaire</span><span class="sxs-lookup"><span data-stu-id="a8837-163">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="a8837-164">(n’enregistrez ***pas*** ce paramètre dans la propriété `parameters`).</span><span class="sxs-lookup"><span data-stu-id="a8837-164">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="a8837-165">Dans le corps de la fonction, un gestionnaire doit être attribué au membre `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="a8837-165">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="a8837-166">Pour plus d’informations, voir [Annuler une fonction](custom-functions-web-reqs.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="a8837-166">For more information, see [Canceling a function](custom-functions-web-reqs.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="a8837-167">boolean</span><span class="sxs-lookup"><span data-stu-id="a8837-167">boolean</span></span>  |  <span data-ttu-id="a8837-168">Non</span><span class="sxs-lookup"><span data-stu-id="a8837-168">No</span></span><br/><br/><span data-ttu-id="a8837-169">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="a8837-169">Default value is `false`.</span></span>  |  <span data-ttu-id="a8837-170">Si la valeur est `true`, la fonction peut envoyer une sortie à la cellule à plusieurs reprises, même en cas d’appel unique.</span><span class="sxs-lookup"><span data-stu-id="a8837-170">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="a8837-171">Cette option est utile pour des sources de données qui changent rapidement, telles que des valeurs boursières.</span><span class="sxs-lookup"><span data-stu-id="a8837-171">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="a8837-172">Si vous utilisez cette option, Excel appelle la fonction JavaScript avec un paramètre `caller` supplémentaire</span><span class="sxs-lookup"><span data-stu-id="a8837-172">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="a8837-173">(n’enregistrez ***pas*** ce paramètre dans la propriété `parameters`).</span><span class="sxs-lookup"><span data-stu-id="a8837-173">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="a8837-174">La fonction ne doit pas utiliser d’instruction `return`.</span><span class="sxs-lookup"><span data-stu-id="a8837-174">The function should have no `return` statement.</span></span> <span data-ttu-id="a8837-175">Au lieu de cela, la valeur obtenue est transmise en tant qu’argument de la méthode de rappel `caller.setResult`.</span><span class="sxs-lookup"><span data-stu-id="a8837-175">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="a8837-176">Pour plus d’informations, voir [Diffusion en continu de fonctions](custom-functions-web-reqs.md#streaming-functions).</span><span class="sxs-lookup"><span data-stu-id="a8837-176">For more information, see [Streaming functions](custom-functions-web-reqs.md#streaming-functions).</span></span> |
|  `volatile`  | <span data-ttu-id="a8837-177">boolean</span><span class="sxs-lookup"><span data-stu-id="a8837-177">boolean</span></span> | <span data-ttu-id="a8837-178">Non</span><span class="sxs-lookup"><span data-stu-id="a8837-178">No</span></span> <br/><br/><span data-ttu-id="a8837-179">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="a8837-179">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="a8837-180">Si la valeur est `true`, la fonction est recalculée à chaque recalcul d’Excel, et plus à chaque fois que les valeurs dépendantes de la formules sont modifiées.</span><span class="sxs-lookup"><span data-stu-id="a8837-180">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="a8837-181">Une fonction ne peut pas être à la fois diffusée en continu et volatile.</span><span class="sxs-lookup"><span data-stu-id="a8837-181">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="a8837-182">Si les propriétés `stream` et `volatile` sont toutes les deux définies sur `true`, l’option volatile est ignorée.</span><span class="sxs-lookup"><span data-stu-id="a8837-182">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="a8837-183">paramètres</span><span class="sxs-lookup"><span data-stu-id="a8837-183">parameters</span></span>

<span data-ttu-id="a8837-184">La propriété `parameters` est un tableau d’objets paramètre.</span><span class="sxs-lookup"><span data-stu-id="a8837-184">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="a8837-185">Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="a8837-185">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="a8837-186">Propriété</span><span class="sxs-lookup"><span data-stu-id="a8837-186">Property</span></span>  |  <span data-ttu-id="a8837-187">Type de données</span><span class="sxs-lookup"><span data-stu-id="a8837-187">Data type</span></span>  |  <span data-ttu-id="a8837-188">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="a8837-188">Required</span></span>  |  <span data-ttu-id="a8837-189">Description</span><span class="sxs-lookup"><span data-stu-id="a8837-189">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="a8837-190">string</span><span class="sxs-lookup"><span data-stu-id="a8837-190">string</span></span>  |  <span data-ttu-id="a8837-191">Non</span><span class="sxs-lookup"><span data-stu-id="a8837-191">No</span></span> |  <span data-ttu-id="a8837-192">Description du paramètre.</span><span class="sxs-lookup"><span data-stu-id="a8837-192">A description of the parameter.</span></span> <span data-ttu-id="a8837-193">S’affiche dans intelliSense d’Excel.</span><span class="sxs-lookup"><span data-stu-id="a8837-193">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="a8837-194">string</span><span class="sxs-lookup"><span data-stu-id="a8837-194">string</span></span>  |  <span data-ttu-id="a8837-195">Non</span><span class="sxs-lookup"><span data-stu-id="a8837-195">No</span></span>  |  <span data-ttu-id="a8837-196">Doit être **scalaire** (valeur autre que de tableau) ou **matrice** (tableau bidimensionnel).</span><span class="sxs-lookup"><span data-stu-id="a8837-196">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="a8837-197">string</span><span class="sxs-lookup"><span data-stu-id="a8837-197">string</span></span>  |  <span data-ttu-id="a8837-198">Oui</span><span class="sxs-lookup"><span data-stu-id="a8837-198">Yes</span></span>  |  <span data-ttu-id="a8837-199">Le nom du paramètre.</span><span class="sxs-lookup"><span data-stu-id="a8837-199">The name of the parameter.</span></span> <span data-ttu-id="a8837-200">Ce nom s’affiche dans intelliSense d’Excel.</span><span class="sxs-lookup"><span data-stu-id="a8837-200">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="a8837-201">string</span><span class="sxs-lookup"><span data-stu-id="a8837-201">string</span></span>  |  <span data-ttu-id="a8837-202">Non</span><span class="sxs-lookup"><span data-stu-id="a8837-202">No</span></span>  |  <span data-ttu-id="a8837-203">Type de données du paramètre.</span><span class="sxs-lookup"><span data-stu-id="a8837-203">The data type of the parameter.</span></span> <span data-ttu-id="a8837-204">Peut être **boolean**, **number**, **string** ou **any** qui vous permet d’utiliser n’importe lequel des trois types précédents.</span><span class="sxs-lookup"><span data-stu-id="a8837-204">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="a8837-205">Si cette propriété n’est pas spécifiée, le type de données par défaut est **any**.</span><span class="sxs-lookup"><span data-stu-id="a8837-205">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="a8837-206">boolean</span><span class="sxs-lookup"><span data-stu-id="a8837-206">boolean</span></span> | <span data-ttu-id="a8837-207">Non</span><span class="sxs-lookup"><span data-stu-id="a8837-207">No</span></span> | <span data-ttu-id="a8837-208">Si la valeur est `true`, le paramètre est facultatif.</span><span class="sxs-lookup"><span data-stu-id="a8837-208">If `true`, the parameter is optional.</span></span> |

>[!NOTE]
> <span data-ttu-id="a8837-209">Si la propriété `type` d’un paramètre facultatif n’est pas spécifiée ou est définie sur `any`, vous remarquerez peut-être des problèmes tels que des erreurs de linting dans votre IDE et des paramètres facultatifs non affichés lorsque la fonction est saisie dans une cellule Excel.</span><span class="sxs-lookup"><span data-stu-id="a8837-209">If the `type` property of an optional parameter is either not specified or set to `any`, you may notice issues such as linting errors in your IDE and optional parameters not being displayed when the function is being entered into a cell in Excel.</span></span> <span data-ttu-id="a8837-210">Ces problèmes seront résolus en décembre 2018.</span><span class="sxs-lookup"><span data-stu-id="a8837-210">This is projected to change in December of 2018.</span></span>

## <a name="result"></a><span data-ttu-id="a8837-211">résultat</span><span class="sxs-lookup"><span data-stu-id="a8837-211">result</span></span>

<span data-ttu-id="a8837-212">L’objet `result` définit le type des informations renvoyées par la fonction.</span><span class="sxs-lookup"><span data-stu-id="a8837-212">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="a8837-213">Le tableau suivant répertorie les propriétés de l’objet `result`.</span><span class="sxs-lookup"><span data-stu-id="a8837-213">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="a8837-214">Propriété</span><span class="sxs-lookup"><span data-stu-id="a8837-214">Property</span></span>  |  <span data-ttu-id="a8837-215">Type de données</span><span class="sxs-lookup"><span data-stu-id="a8837-215">Data type</span></span>  |  <span data-ttu-id="a8837-216">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="a8837-216">Required</span></span>  |  <span data-ttu-id="a8837-217">Description</span><span class="sxs-lookup"><span data-stu-id="a8837-217">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="a8837-218">string</span><span class="sxs-lookup"><span data-stu-id="a8837-218">string</span></span>  |  <span data-ttu-id="a8837-219">Non</span><span class="sxs-lookup"><span data-stu-id="a8837-219">No</span></span>  |  <span data-ttu-id="a8837-220">Doit être **scalaire** (valeur autre que de tableau) ou **matrice** (tableau bidimensionnel).</span><span class="sxs-lookup"><span data-stu-id="a8837-220">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="a8837-221">string</span><span class="sxs-lookup"><span data-stu-id="a8837-221">string</span></span>  |  <span data-ttu-id="a8837-222">Oui</span><span class="sxs-lookup"><span data-stu-id="a8837-222">Yes</span></span>  |  <span data-ttu-id="a8837-223">Type de données du paramètre.</span><span class="sxs-lookup"><span data-stu-id="a8837-223">The data type of the parameter.</span></span> <span data-ttu-id="a8837-224">Doit être **boolean**, **number**, **string** ou **any** qui vous permet d’utiliser n’importe lequel des trois types précédents.</span><span class="sxs-lookup"><span data-stu-id="a8837-224">Must be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> |

## <a name="see-also"></a><span data-ttu-id="a8837-225">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a8837-225">See also</span></span>

* [<span data-ttu-id="a8837-226">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="a8837-226">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="a8837-227">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="a8837-227">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="a8837-228">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="a8837-228">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="a8837-229">Fonctions personnalisées changelog</span><span class="sxs-lookup"><span data-stu-id="a8837-229">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="a8837-230">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="a8837-230">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
