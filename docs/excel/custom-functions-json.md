---
ms.date: 10/17/2018
description: Définir les métadonnées pour des fonctions personnalisées dans Excel.
title: Métadonnées pour des fonctions personnalisées dans Excel
ms.openlocfilehash: cff1cbc22f39c99597d4abe7005d7b8bbce6e185
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640007"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="ce726-103">Métadonnées des fonctions personnalisées (aperçu)</span><span class="sxs-lookup"><span data-stu-id="ce726-103">Custom functions metadata</span></span>

<span data-ttu-id="ce726-p101">Lorsque vous définissez des [fonctions personnalisées](custom-functions-overview.md) dans votre complément Excel, votre projet de complément doit inclure un fichier de métadonnées JSON qui fournit les informations nécessaires pour enregistrer les fonctions personnalisées et les rendre disponibles pour les utilisateurs finaux dans Excel. Cet article décrit le format du fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="ce726-p101">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users. This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="ce726-106">Pour plus d’informations sur les autres fichiers que vous devez inclure dans votre projet de complément pour activer les fonctions personnalisées, voir [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="ce726-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="ce726-107">Métadonnées d’exemple</span><span class="sxs-lookup"><span data-stu-id="ce726-107">Example metadata</span></span>

<span data-ttu-id="ce726-p102">L’exemple suivant montre le contenu d’un fichier de métadonnées JSON pour un complément qui définit les fonctions personnalisées. Les sections qui suivent cet exemple fournissent des informations détaillées sur les propriétés individuelles dans cet exemple JSON.</span><span class="sxs-lookup"><span data-stu-id="ce726-p102">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions. The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="ce726-110">Un fichier d’exemple JSON complet est disponible dans le référentiel GitHub [ OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="ce726-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions GitHub repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions"></a><span data-ttu-id="ce726-111">fonctions</span><span class="sxs-lookup"><span data-stu-id="ce726-111">functions</span></span> 

<span data-ttu-id="ce726-p103">La propriété `functions` est un tableau d’objets de fonctions personnalisées.. Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="ce726-p103">The `functions` property is an array of custom function objects. The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="ce726-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="ce726-114">Property</span></span>  |  <span data-ttu-id="ce726-115">Type de données</span><span class="sxs-lookup"><span data-stu-id="ce726-115">Data type</span></span>  |  <span data-ttu-id="ce726-116">Requis</span><span class="sxs-lookup"><span data-stu-id="ce726-116">Required</span></span>  |  <span data-ttu-id="ce726-117">Description</span><span class="sxs-lookup"><span data-stu-id="ce726-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="ce726-118">string</span><span class="sxs-lookup"><span data-stu-id="ce726-118">string</span></span>  |  <span data-ttu-id="ce726-119">Non</span><span class="sxs-lookup"><span data-stu-id="ce726-119">No</span></span>  |  <span data-ttu-id="ce726-p104">Description de la fonction que les utilisateurs voient dans Excel. Par exemple, **Convertit une valeur en Celsius en Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="ce726-p104">The description of the function that end users see in Excel. For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="ce726-122">string</span><span class="sxs-lookup"><span data-stu-id="ce726-122">string</span></span>  |   <span data-ttu-id="ce726-123">Non</span><span class="sxs-lookup"><span data-stu-id="ce726-123">No</span></span>  |  <span data-ttu-id="ce726-p105">URL qui fournit des informations sur la fonction. (Elle est affichée dans un volet Office.) Par exemple, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="ce726-p105">URL that provides information about the function. (It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="ce726-126">string</span><span class="sxs-lookup"><span data-stu-id="ce726-126">string</span></span> | <span data-ttu-id="ce726-127">Oui</span><span class="sxs-lookup"><span data-stu-id="ce726-127">Yes</span></span> | <span data-ttu-id="ce726-p106">ID unique de la fonction. Ce code ne peut contenir que des caractères alphanumériques et des points et ne doit pas être modifié après sa définition.</span><span class="sxs-lookup"><span data-stu-id="ce726-p106">A unique ID for the function. This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="ce726-130">string</span><span class="sxs-lookup"><span data-stu-id="ce726-130">string</span></span>  |  <span data-ttu-id="ce726-131">Oui</span><span class="sxs-lookup"><span data-stu-id="ce726-131">Yes</span></span>  |  <span data-ttu-id="ce726-p107">Nom de la fonction que l’utilisateur final voit dans Excel. Dans Excel, ce nom de fonction aura pour préfixe l’espace de noms des fonctions personnalisées qui est spécifié dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="ce726-p107">The name of the function that end users see in Excel. In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="ce726-134">objet</span><span class="sxs-lookup"><span data-stu-id="ce726-134">object</span></span>  |  <span data-ttu-id="ce726-135">Non</span><span class="sxs-lookup"><span data-stu-id="ce726-135">No</span></span>  |  <span data-ttu-id="ce726-p108">Permet de personnaliser en partie comment et quand Excel exécute la fonction. Voir l' [objet options](#options-object) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="ce726-p108">Enables you to customize some aspects of how and when Excel executes the function. See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="ce726-138">array</span><span class="sxs-lookup"><span data-stu-id="ce726-138">array</span></span>  |  <span data-ttu-id="ce726-139">Oui</span><span class="sxs-lookup"><span data-stu-id="ce726-139">Yes</span></span>  |  <span data-ttu-id="ce726-p109">Tableau qui définit les paramètres d’entrée de la fonction. Consultez [Tableau de paramètres](#parameters-array) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="ce726-p109">Array that defines the input parameters for the function. See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="ce726-142">object</span><span class="sxs-lookup"><span data-stu-id="ce726-142">object</span></span>  |  <span data-ttu-id="ce726-143">Oui</span><span class="sxs-lookup"><span data-stu-id="ce726-143">Yes</span></span>  |  <span data-ttu-id="ce726-p110">Objet qui définit le type d’informations renvoyées par la fonction. Voir l' [Objet de résultat](#result-object) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="ce726-p110">Object that defines the type of information that is returned by the function. See [result object](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="ce726-146">options</span><span class="sxs-lookup"><span data-stu-id="ce726-146">options</span></span>

<span data-ttu-id="ce726-p111">L’objet `options` vous permet de personnaliser en partie comment et quand Excel exécute la fonction. Le tableau suivant répertorie les propriétés de l'objet  `options`.</span><span class="sxs-lookup"><span data-stu-id="ce726-p111">The `options` object enables you to customize some aspects of how and when Excel executes the function. The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="ce726-149">Propriété</span><span class="sxs-lookup"><span data-stu-id="ce726-149">Property</span></span>  |  <span data-ttu-id="ce726-150">Type de données</span><span class="sxs-lookup"><span data-stu-id="ce726-150">Data type</span></span>  |  <span data-ttu-id="ce726-151">Requis</span><span class="sxs-lookup"><span data-stu-id="ce726-151">Required</span></span>  |  <span data-ttu-id="ce726-152">Description</span><span class="sxs-lookup"><span data-stu-id="ce726-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="ce726-153">boolean</span><span class="sxs-lookup"><span data-stu-id="ce726-153">boolean</span></span>  |  <span data-ttu-id="ce726-154">Non</span><span class="sxs-lookup"><span data-stu-id="ce726-154">No</span></span><br/><br/><span data-ttu-id="ce726-155">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="ce726-155">Default value is 4.</span></span>  |  <span data-ttu-id="ce726-p112">Si `true`, Excel appelle le gestionnaire `onCanceled` à chaque fois que l’utilisateur exécute une action qui a pour effet l’annulation de la fonction ; par exemple, déclencher manuellement le recalcul, ou modifier une cellule référencée par la fonction. Si vous utilisez cette option, Excel appelle la fonction JavaScript avec un paramètre `caller` supplémentaire. (Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`). Dans le corps de la fonction, un gestionnaire doit être affecté au membre `caller.onCanceled`. Pour plus d’informations, voir [Annulation d’une fonction](custom-functions-overview.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="ce726-p112">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). In the body of the function, a handler must be assigned to the `caller.onCanceled` member. For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="ce726-161">boolean</span><span class="sxs-lookup"><span data-stu-id="ce726-161">boolean</span></span>  |  <span data-ttu-id="ce726-162">Non</span><span class="sxs-lookup"><span data-stu-id="ce726-162">No</span></span><br/><br/><span data-ttu-id="ce726-163">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="ce726-163">Default value is 4.</span></span>  |  <span data-ttu-id="ce726-p113">Si `true`, la fonction peut déclencher le recalcul d'une cellule de manière répétée, même lorsqu’elle est appelée une seule fois. Cette option est utile pour les sources de données qui évoluent rapidement, telles que des actions. Si vous utilisez cette option, Excel appelle la fonction JavaScript avec un paramètre `caller` supplémentaire. (Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`). La fonction ne devrait pas avoir de déclaration `return`. Au lieu de cela, la valeur du résultat est transmise en tant que motif de la méthode de rappel  `caller.setResult`. Pour plus d’informations, voir [Diffusion en continu d’une fonction](custom-functions-overview.md#streaming-functions).</span><span class="sxs-lookup"><span data-stu-id="ce726-p113">If `true`, the function can output repeatedly to the cell even when invoked only once. This option is useful for rapidly-changing data sources, such as a stock price. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). The function should have no `return` statement. Instead, the result value is passed as the argument of the `caller.setResult` callback method. For more information, see [Streaming functions](custom-functions-overview.md#streaming-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="ce726-171">parameters</span><span class="sxs-lookup"><span data-stu-id="ce726-171">parameters</span></span>

<span data-ttu-id="ce726-p114">La propriété `parameters` est un tableau de paramètres d'objets. Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="ce726-p114">The `parameters` property is an array of parameter objects. The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="ce726-174">Propriété</span><span class="sxs-lookup"><span data-stu-id="ce726-174">Property</span></span>  |  <span data-ttu-id="ce726-175">Type de données</span><span class="sxs-lookup"><span data-stu-id="ce726-175">Data type</span></span>  |  <span data-ttu-id="ce726-176">Requis</span><span class="sxs-lookup"><span data-stu-id="ce726-176">Required</span></span>  |  <span data-ttu-id="ce726-177">Description</span><span class="sxs-lookup"><span data-stu-id="ce726-177">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="ce726-178">string</span><span class="sxs-lookup"><span data-stu-id="ce726-178">string</span></span>  |  <span data-ttu-id="ce726-179">Non</span><span class="sxs-lookup"><span data-stu-id="ce726-179">No</span></span> |  <span data-ttu-id="ce726-180">Description du paramètre.</span><span class="sxs-lookup"><span data-stu-id="ce726-180">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="ce726-181">string</span><span class="sxs-lookup"><span data-stu-id="ce726-181">string</span></span>  |  <span data-ttu-id="ce726-182">Non</span><span class="sxs-lookup"><span data-stu-id="ce726-182">No</span></span>  |  <span data-ttu-id="ce726-183">Doit être **scalar** (une valeur non tableau) ou **matrix** (tableau à deux dimensions).</span><span class="sxs-lookup"><span data-stu-id="ce726-183">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="ce726-184">string</span><span class="sxs-lookup"><span data-stu-id="ce726-184">string</span></span>  |  <span data-ttu-id="ce726-185">Oui</span><span class="sxs-lookup"><span data-stu-id="ce726-185">Yes</span></span>  |  <span data-ttu-id="ce726-p115">Le nom du paramètre. Ce nom est affiché dans intelliSense d'Excel.</span><span class="sxs-lookup"><span data-stu-id="ce726-p115">The name of the parameter. This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="ce726-188">string</span><span class="sxs-lookup"><span data-stu-id="ce726-188">string</span></span>  |  <span data-ttu-id="ce726-189">Non</span><span class="sxs-lookup"><span data-stu-id="ce726-189">No</span></span>  |  <span data-ttu-id="ce726-p116">Le type de données du paramètre. Doit être **boolean**, **number**ou **string**.</span><span class="sxs-lookup"><span data-stu-id="ce726-p116">The data type of the parameter. Must be **boolean**, **number**, or **string**.</span></span>  |

## <a name="result"></a><span data-ttu-id="ce726-192">result</span><span class="sxs-lookup"><span data-stu-id="ce726-192">result</span></span>

<span data-ttu-id="ce726-p117">L'objet `results` définit le type d’informations renvoyées par la fonction. Le tableau suivant répertorie les propriétés de l'objet `result` .</span><span class="sxs-lookup"><span data-stu-id="ce726-p117">The `results` object defines the type of information that is returned by the function. The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="ce726-195">Propriété</span><span class="sxs-lookup"><span data-stu-id="ce726-195">Property</span></span>  |  <span data-ttu-id="ce726-196">Type de données</span><span class="sxs-lookup"><span data-stu-id="ce726-196">Data type</span></span>  |  <span data-ttu-id="ce726-197">Requis</span><span class="sxs-lookup"><span data-stu-id="ce726-197">Required</span></span>  |  <span data-ttu-id="ce726-198">Description</span><span class="sxs-lookup"><span data-stu-id="ce726-198">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="ce726-199">string</span><span class="sxs-lookup"><span data-stu-id="ce726-199">string</span></span>  |  <span data-ttu-id="ce726-200">Non</span><span class="sxs-lookup"><span data-stu-id="ce726-200">No</span></span>  |  <span data-ttu-id="ce726-201">Doit être **scalar** (une valeur non tableau) ou **matrix** (tableau à deux dimensions).</span><span class="sxs-lookup"><span data-stu-id="ce726-201">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="ce726-202">string</span><span class="sxs-lookup"><span data-stu-id="ce726-202">string</span></span>  |  <span data-ttu-id="ce726-203">Oui</span><span class="sxs-lookup"><span data-stu-id="ce726-203">Yes</span></span>  |  <span data-ttu-id="ce726-p118">Le type de données du paramètre. Doit être **boolean**, **number**ou **string**.</span><span class="sxs-lookup"><span data-stu-id="ce726-p118">The data type of the parameter. Must be **boolean**, **number**, or **string**.</span></span>  |

## <a name="see-also"></a><span data-ttu-id="ce726-206">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ce726-206">See also</span></span>

* [<span data-ttu-id="ce726-207">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="ce726-207">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="ce726-208">Runtime pour les fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="ce726-208">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="ce726-209">Meilleures pratiques pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="ce726-209">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="ce726-210">Didacticiel sur les fonctions personnalisées d’Excel</span><span class="sxs-lookup"><span data-stu-id="ce726-210">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
