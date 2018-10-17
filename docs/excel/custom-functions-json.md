---
ms.date: 09/27/2018
description: Définir les métadonnées pour des fonctions personnalisées dans Excel.
title: Métadonnées pour des fonctions personnalisées dans Excel
ms.openlocfilehash: b7c7f26d56309f43758b9b13025ebaad661aeaed
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579869"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="a4cf3-103">Métadonnées des fonctions personnalisées (aperçu)</span><span class="sxs-lookup"><span data-stu-id="a4cf3-103">Custom functions metadata</span></span>

<span data-ttu-id="a4cf3-p101">Lorsque vous définissez des [fonctions personnalisées](custom-functions-overview.md) dans votre complément Excel, votre projet de complément doit inclure un fichier de métadonnées JSON qui fournit les informations nécessaires pour enregistrer les fonctions personnalisées et les rendre disponibles pour les utilisateurs finaux dans Excel. Cet article décrit le format du fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-p101">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users. This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="a4cf3-106">Pour plus d’informations sur les autres fichiers que vous devez inclure dans votre projet de complément pour activer les fonctions personnalisées, voir [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="a4cf3-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="a4cf3-107">Métadonnées d’exemple</span><span class="sxs-lookup"><span data-stu-id="a4cf3-107">Example metadata</span></span>

<span data-ttu-id="a4cf3-p102">L’exemple suivant montre le contenu d’un fichier de métadonnées JSON pour un complément qui définit les fonctions personnalisées. Les sections qui suivent cet exemple fournissent des informations détaillées sur les propriétés individuelles dans cet exemple JSON.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-p102">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions. The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="a4cf3-110">Un fichier d’exemple JSON complet est disponible dans le référentiel GitHub [ OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="a4cf3-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions GitHub repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions"></a><span data-ttu-id="a4cf3-111">fonctions</span><span class="sxs-lookup"><span data-stu-id="a4cf3-111">functions</span></span> 

<span data-ttu-id="a4cf3-p103">La propriété `functions` est un tableau d’objets de fonctions personnalisées.. Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-p103">The `functions` property is an array of custom function objects. The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="a4cf3-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="a4cf3-114">Property</span></span>  |  <span data-ttu-id="a4cf3-115">Type de données</span><span class="sxs-lookup"><span data-stu-id="a4cf3-115">Data type</span></span>  |  <span data-ttu-id="a4cf3-116">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="a4cf3-116">Required</span></span>  |  <span data-ttu-id="a4cf3-117">Description</span><span class="sxs-lookup"><span data-stu-id="a4cf3-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="a4cf3-118">string</span><span class="sxs-lookup"><span data-stu-id="a4cf3-118">string</span></span>  |  <span data-ttu-id="a4cf3-119">Non</span><span class="sxs-lookup"><span data-stu-id="a4cf3-119">No</span></span>  |  <span data-ttu-id="a4cf3-p104">Description de la fonction que les utilisateurs voient dans Excel. Par exemple, **Convertit une valeur en Celsius en Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-p104">The description of the function that end users see in Excel. For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="a4cf3-122">string</span><span class="sxs-lookup"><span data-stu-id="a4cf3-122">string</span></span>  |   <span data-ttu-id="a4cf3-123">Non</span><span class="sxs-lookup"><span data-stu-id="a4cf3-123">No</span></span>  |  <span data-ttu-id="a4cf3-p105">URL qui fournit des informations sur la fonction. (Elle est affichée dans un volet Office.) Par exemple, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-p105">URL that provides information about the function. (It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="a4cf3-126">string</span><span class="sxs-lookup"><span data-stu-id="a4cf3-126">string</span></span> | <span data-ttu-id="a4cf3-127">Oui</span><span class="sxs-lookup"><span data-stu-id="a4cf3-127">Yes</span></span> | <span data-ttu-id="a4cf3-128">ID unique de la fonction.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-128">A unique ID for the function.</span></span> <span data-ttu-id="a4cf3-129">Ce code ne peut contenir que des caractères alphanumériques et des périodes et ne doit pas être modifié après sa définition.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-129">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="a4cf3-130">string</span><span class="sxs-lookup"><span data-stu-id="a4cf3-130">string</span></span>  |  <span data-ttu-id="a4cf3-131">Oui</span><span class="sxs-lookup"><span data-stu-id="a4cf3-131">Yes</span></span>  |  <span data-ttu-id="a4cf3-p107">Nom de la fonction que l’utilisateur final voit dans Excel. Dans Excel, ce nom de fonction aura pour préfixe l’espace de noms des fonctions personnalisées qui est spécifié dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-p107">The name of the function that end users see in Excel. In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="a4cf3-134">objet</span><span class="sxs-lookup"><span data-stu-id="a4cf3-134">object</span></span>  |  <span data-ttu-id="a4cf3-135">Non</span><span class="sxs-lookup"><span data-stu-id="a4cf3-135">No</span></span>  |  <span data-ttu-id="a4cf3-p108">Permet de personnaliser en partie comment et quand Excel exécute la fonction. Voir l' [objet options](#options-object) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-p108">Enables you to customize some aspects of how and when Excel executes the function. See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="a4cf3-138">array</span><span class="sxs-lookup"><span data-stu-id="a4cf3-138">array</span></span>  |  <span data-ttu-id="a4cf3-139">Oui</span><span class="sxs-lookup"><span data-stu-id="a4cf3-139">Yes</span></span>  |  <span data-ttu-id="a4cf3-p109">Tableau qui définit les paramètres d’entrée de la fonction. Consultez [Tableau de paramètres](#parameters-array) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-p109">Array that defines the input parameters for the function. See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="a4cf3-142">objet</span><span class="sxs-lookup"><span data-stu-id="a4cf3-142">object</span></span>  |  <span data-ttu-id="a4cf3-143">Oui</span><span class="sxs-lookup"><span data-stu-id="a4cf3-143">Yes</span></span>  |  <span data-ttu-id="a4cf3-p110">Objet qui définit le type d’informations renvoyées par la fonction. Voir l' [Objet de résultat](#result-object) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-p110">Object that defines the type of information that is returned by the function. See [result object](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="a4cf3-146">options</span><span class="sxs-lookup"><span data-stu-id="a4cf3-146">options</span></span>

<span data-ttu-id="a4cf3-p111">L’objet `options` vous permet de personnaliser en partie comment et quand Excel exécute la fonction. Le tableau suivant répertorie les propriétés de l'objet  `options`.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-p111">The `options` object enables you to customize some aspects of how and when Excel executes the function. The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="a4cf3-149">Propriété</span><span class="sxs-lookup"><span data-stu-id="a4cf3-149">Property</span></span>  |  <span data-ttu-id="a4cf3-150">Type de données</span><span class="sxs-lookup"><span data-stu-id="a4cf3-150">Data type</span></span>  |  <span data-ttu-id="a4cf3-151">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="a4cf3-151">Required</span></span>  |  <span data-ttu-id="a4cf3-152">Description</span><span class="sxs-lookup"><span data-stu-id="a4cf3-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="a4cf3-153">boolean</span><span class="sxs-lookup"><span data-stu-id="a4cf3-153">boolean</span></span>  |  <span data-ttu-id="a4cf3-154">Non</span><span class="sxs-lookup"><span data-stu-id="a4cf3-154">No</span></span><br/><br/><span data-ttu-id="a4cf3-155">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-155">Default value is 4.</span></span>  |  <span data-ttu-id="a4cf3-p112">Si `true`, Excel appelle le gestionnaire `onCanceled` à chaque fois que l’utilisateur exécute une action qui a pour effet l’annulation de la fonction ; par exemple, déclencher manuellement le recalcul, ou modifier une cellule référencée par la fonction. Si vous utilisez cette option, Excel appelle la fonction JavaScript avec un paramètre `caller` supplémentaire. (Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`). Dans le corps de la fonction, un gestionnaire doit être affecté au membre `caller.onCanceled`. Pour plus d’informations, voir [Annulation d’une fonction](custom-functions-overview.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="a4cf3-p112">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). In the body of the function, a handler must be assigned to the `caller.onCanceled` member. For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="a4cf3-161">boolean</span><span class="sxs-lookup"><span data-stu-id="a4cf3-161">boolean</span></span>  |  <span data-ttu-id="a4cf3-162">Non</span><span class="sxs-lookup"><span data-stu-id="a4cf3-162">No</span></span><br/><br/><span data-ttu-id="a4cf3-163">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-163">Default value is 4.</span></span>  |  <span data-ttu-id="a4cf3-164">Si `true`, la fonction peut générer une sortie plusieurs fois dans la cellule même lorsqu’elle n’est invoquée qu’une seule fois.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-164">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="a4cf3-165">Cette option est utile pour les sources de données en évolution rapide, telles que le cours d'une action.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-165">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="a4cf3-166">Si vous utilisez cette option, Excel appellera la fonction JavaScript avec un paramètre `caller` additionnel.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-166">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="a4cf3-167">(Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`).</span><span class="sxs-lookup"><span data-stu-id="a4cf3-167">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="a4cf3-168">La fonction ne devrait pas avoir d’instruction `return`.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-168">The function should have no `return` statement.</span></span> <span data-ttu-id="a4cf3-169">Au lieu de cela, la valeur du résultat est transmise en tant que motif de la `caller.setResult` méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-169">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="a4cf3-170">Pour plus d’informations, voir [les fonctions de diffusion en continu](custom-functions-overview.md#streaming-functions).</span><span class="sxs-lookup"><span data-stu-id="a4cf3-170">For more information, see [Streaming functions](custom-functions-overview.md#streaming-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="a4cf3-171">parameters</span><span class="sxs-lookup"><span data-stu-id="a4cf3-171">parameters</span></span>

<span data-ttu-id="a4cf3-p114">La propriété `parameters` est un tableau de paramètres d'objets. Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-p114">The `parameters` property is an array of parameter objects. The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="a4cf3-174">Propriété</span><span class="sxs-lookup"><span data-stu-id="a4cf3-174">Property</span></span>  |  <span data-ttu-id="a4cf3-175">Type de données</span><span class="sxs-lookup"><span data-stu-id="a4cf3-175">Data type</span></span>  |  <span data-ttu-id="a4cf3-176">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="a4cf3-176">Required</span></span>  |  <span data-ttu-id="a4cf3-177">Description</span><span class="sxs-lookup"><span data-stu-id="a4cf3-177">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="a4cf3-178">string</span><span class="sxs-lookup"><span data-stu-id="a4cf3-178">string</span></span>  |  <span data-ttu-id="a4cf3-179">Non</span><span class="sxs-lookup"><span data-stu-id="a4cf3-179">No</span></span> |  <span data-ttu-id="a4cf3-180">Description du paramètre.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-180">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="a4cf3-181">string</span><span class="sxs-lookup"><span data-stu-id="a4cf3-181">string</span></span>  |  <span data-ttu-id="a4cf3-182">Non</span><span class="sxs-lookup"><span data-stu-id="a4cf3-182">No</span></span>  |  <span data-ttu-id="a4cf3-183">Doit être **scalar** (une valeur non tableau) ou **matrix** (tableau à deux dimensions).</span><span class="sxs-lookup"><span data-stu-id="a4cf3-183">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="a4cf3-184">string</span><span class="sxs-lookup"><span data-stu-id="a4cf3-184">string</span></span>  |  <span data-ttu-id="a4cf3-185">Oui</span><span class="sxs-lookup"><span data-stu-id="a4cf3-185">Yes</span></span>  |  <span data-ttu-id="a4cf3-p115">Le nom du paramètre. Ce nom est affiché dans intelliSense d'Excel.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-p115">The name of the parameter. This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="a4cf3-188">string</span><span class="sxs-lookup"><span data-stu-id="a4cf3-188">string</span></span>  |  <span data-ttu-id="a4cf3-189">Non</span><span class="sxs-lookup"><span data-stu-id="a4cf3-189">No</span></span>  |  <span data-ttu-id="a4cf3-p116">Le type de données du paramètre. Doit être **boolean**, **number**ou **string**.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-p116">The data type of the parameter. Must be **boolean**, **number**, or **string**.</span></span>  |

## <a name="result"></a><span data-ttu-id="a4cf3-192">result</span><span class="sxs-lookup"><span data-stu-id="a4cf3-192">result</span></span>

<span data-ttu-id="a4cf3-p117">L'objet `results` définit le type d’informations renvoyées par la fonction. Le tableau suivant répertorie les propriétés de l'objet `result` .</span><span class="sxs-lookup"><span data-stu-id="a4cf3-p117">The `results` object defines the type of information that is returned by the function. The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="a4cf3-195">Propriété</span><span class="sxs-lookup"><span data-stu-id="a4cf3-195">Property</span></span>  |  <span data-ttu-id="a4cf3-196">Type de données</span><span class="sxs-lookup"><span data-stu-id="a4cf3-196">Data type</span></span>  |  <span data-ttu-id="a4cf3-197">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="a4cf3-197">Required</span></span>  |  <span data-ttu-id="a4cf3-198">Description</span><span class="sxs-lookup"><span data-stu-id="a4cf3-198">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="a4cf3-199">string</span><span class="sxs-lookup"><span data-stu-id="a4cf3-199">string</span></span>  |  <span data-ttu-id="a4cf3-200">Non</span><span class="sxs-lookup"><span data-stu-id="a4cf3-200">No</span></span>  |  <span data-ttu-id="a4cf3-201">Doit être **scalar** (une valeur non tableau) ou **matrix** (tableau à deux dimensions).</span><span class="sxs-lookup"><span data-stu-id="a4cf3-201">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="a4cf3-202">string</span><span class="sxs-lookup"><span data-stu-id="a4cf3-202">string</span></span>  |  <span data-ttu-id="a4cf3-203">Oui</span><span class="sxs-lookup"><span data-stu-id="a4cf3-203">Yes</span></span>  |  <span data-ttu-id="a4cf3-p118">Le type de données du paramètre. Doit être **boolean**, **number**ou **string**.</span><span class="sxs-lookup"><span data-stu-id="a4cf3-p118">The data type of the parameter. Must be **boolean**, **number**, or **string**.</span></span>  |

## <a name="see-also"></a><span data-ttu-id="a4cf3-206">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a4cf3-206">See also</span></span>

* [<span data-ttu-id="a4cf3-207">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="a4cf3-207">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="a4cf3-208">Runtime pour les fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="a4cf3-208">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="a4cf3-209">Meilleures pratiques pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="a4cf3-209">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="a4cf3-210">Didacticiel sur les fonctions personnalisées d’Excel</span><span class="sxs-lookup"><span data-stu-id="a4cf3-210">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)