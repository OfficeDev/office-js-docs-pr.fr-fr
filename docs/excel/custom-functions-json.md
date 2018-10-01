---
ms.date: 09/27/2018
description: Définir les métadonnées pour des fonctions personnalisées dans Excel.
title: Métadonnées pour des fonctions personnalisées dans Excel
ms.openlocfilehash: a179a9c4bc071200cab1377c5e48913bfc8358cf
ms.sourcegitcommit: 1852ae367de53deb91d03ca55d16eb69709340d3
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/29/2018
ms.locfileid: "25348793"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="d3c19-103">Métadonnées des fonctions personnalisées (aperçu)</span><span class="sxs-lookup"><span data-stu-id="d3c19-103">Custom functions metadata</span></span>

<span data-ttu-id="d3c19-p101">Lorsque vous définissez des [fonctions personnalisées](custom-functions-overview.md) dans votre complément Excel, votre projet de complément doit inclure un fichier de métadonnées JSON qui fournit les informations nécessaires pour enregistrer les fonctions personnalisées et les rendre disponibles pour les utilisateurs finaux dans Excel. Cet article décrit le format du fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="d3c19-p101">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end-users. This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="d3c19-106">Pour plus d’informations sur les autres fichiers que vous devez inclure dans votre projet de complément pour activer les fonctions personnalisées, voir [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="d3c19-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="d3c19-107">Métadonnées d’exemple</span><span class="sxs-lookup"><span data-stu-id="d3c19-107">Example metadata</span></span>

<span data-ttu-id="d3c19-108">L’exemple suivant montre le contenu d’un fichier de métadonnées JSON pour un complément qui définit des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="d3c19-108">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="d3c19-109">Les sections qui suivent cet exemple fournissent des informations détaillées sur les propriétés individuelles dans cet exemple JSON.</span><span class="sxs-lookup"><span data-stu-id="d3c19-109">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="d3c19-110">Un fichier d’exemple JSON complet est disponible dans le [référentiel GitHub OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="d3c19-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions GitHub repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions"></a><span data-ttu-id="d3c19-111">functions</span><span class="sxs-lookup"><span data-stu-id="d3c19-111">functions</span></span> 

<span data-ttu-id="d3c19-112">La propriété `functions` est un tableau d’objets de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="d3c19-112">The `functions` property is an array of objects.</span></span> <span data-ttu-id="d3c19-113">Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="d3c19-113">The following table lists the properties of the SP.FieldRatingScale object.</span></span>

|  <span data-ttu-id="d3c19-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="d3c19-114">Property</span></span>  |  <span data-ttu-id="d3c19-115">Type de données</span><span class="sxs-lookup"><span data-stu-id="d3c19-115">Data type</span></span>  |  <span data-ttu-id="d3c19-116">Requis</span><span class="sxs-lookup"><span data-stu-id="d3c19-116">Required</span></span>  |  <span data-ttu-id="d3c19-117">Description</span><span class="sxs-lookup"><span data-stu-id="d3c19-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="d3c19-118">chaîne</span><span class="sxs-lookup"><span data-stu-id="d3c19-118">string</span></span>  |  <span data-ttu-id="d3c19-119">Non</span><span class="sxs-lookup"><span data-stu-id="d3c19-119">No</span></span>  |  <span data-ttu-id="d3c19-p104">Description de la fonction que les utilisateurs voient dans Excel. Par exemple, **Convertit une valeur en Celsius en Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="d3c19-p104">A description of the function that appears in the Excel UI. For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="d3c19-122">chaîne</span><span class="sxs-lookup"><span data-stu-id="d3c19-122">string</span></span>  |   <span data-ttu-id="d3c19-123">Non</span><span class="sxs-lookup"><span data-stu-id="d3c19-123">No</span></span>  |  <span data-ttu-id="d3c19-p105">URL qui fournit des informations sur la fonction. (Elle est affichée dans un volet Office.) Par exemple, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="d3c19-p105">URL where users can get information about the function. (It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="d3c19-126">chaîne</span><span class="sxs-lookup"><span data-stu-id="d3c19-126">string</span></span> | <span data-ttu-id="d3c19-127">Oui</span><span class="sxs-lookup"><span data-stu-id="d3c19-127">Yes</span></span> | <span data-ttu-id="d3c19-128">ID unique de la fonction.</span><span class="sxs-lookup"><span data-stu-id="d3c19-128">A unique ID for the group.</span></span> <span data-ttu-id="d3c19-129">Cet ID ne doit pas être modifié après sa définition.</span><span class="sxs-lookup"><span data-stu-id="d3c19-129">This ID should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="d3c19-130">chaîne</span><span class="sxs-lookup"><span data-stu-id="d3c19-130">string</span></span>  |  <span data-ttu-id="d3c19-131">Oui</span><span class="sxs-lookup"><span data-stu-id="d3c19-131">Yes</span></span>  |  <span data-ttu-id="d3c19-132">Nom de la fonction que les utilisateurs voient dans Excel.</span><span class="sxs-lookup"><span data-stu-id="d3c19-132">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="d3c19-133">Dans Excel, ce nom de fonction sera préfixé par l’espace de noms des fonctions personnalisées spécifié dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="d3c19-133">In the autocomplete menu, this value will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="d3c19-134">object</span><span class="sxs-lookup"><span data-stu-id="d3c19-134">object</span></span>  |  <span data-ttu-id="d3c19-135">Non</span><span class="sxs-lookup"><span data-stu-id="d3c19-135">No</span></span>  |  <span data-ttu-id="d3c19-136">Vous permet de personnaliser certains aspects de la façon dont Excel exécute la fonction, et quand.</span><span class="sxs-lookup"><span data-stu-id="d3c19-136">The  property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="d3c19-137">Voir [objet options](#options-object) pour plus de détails.</span><span class="sxs-lookup"><span data-stu-id="d3c19-137">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="d3c19-138">array</span><span class="sxs-lookup"><span data-stu-id="d3c19-138">array</span></span>  |  <span data-ttu-id="d3c19-139">Oui</span><span class="sxs-lookup"><span data-stu-id="d3c19-139">Yes</span></span>  |  <span data-ttu-id="d3c19-140">Tableau qui définit les paramètres d’entrée de la fonction.</span><span class="sxs-lookup"><span data-stu-id="d3c19-140">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="d3c19-141">Voir[tableau parameters](#parameters-array) pour plus de détails.</span><span class="sxs-lookup"><span data-stu-id="d3c19-141">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="d3c19-142">object</span><span class="sxs-lookup"><span data-stu-id="d3c19-142">object</span></span>  |  <span data-ttu-id="d3c19-143">Oui</span><span class="sxs-lookup"><span data-stu-id="d3c19-143">Yes</span></span>  |  <span data-ttu-id="d3c19-144">Objet qui définit le type de l’information renvoyée par la fonction.</span><span class="sxs-lookup"><span data-stu-id="d3c19-144">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="d3c19-145">Voir [objet result](#result-object) pour plus de détails.</span><span class="sxs-lookup"><span data-stu-id="d3c19-145">See [result object](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="d3c19-146">options</span><span class="sxs-lookup"><span data-stu-id="d3c19-146">options</span></span>

<span data-ttu-id="d3c19-147">L’objet `options` vous permet de personnaliser certains aspects de la façon dont Excel exécute la fonction, et quand.</span><span class="sxs-lookup"><span data-stu-id="d3c19-147">The `options` property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="d3c19-148">Le tableau suivant répertorie les propriétés de l’objet `options`.</span><span class="sxs-lookup"><span data-stu-id="d3c19-148">The following table lists the properties of the</span></span>

|  <span data-ttu-id="d3c19-149">Propriété</span><span class="sxs-lookup"><span data-stu-id="d3c19-149">Property</span></span>  |  <span data-ttu-id="d3c19-150">Type de données</span><span class="sxs-lookup"><span data-stu-id="d3c19-150">Data type</span></span>  |  <span data-ttu-id="d3c19-151">Requis</span><span class="sxs-lookup"><span data-stu-id="d3c19-151">Required</span></span>  |  <span data-ttu-id="d3c19-152">Description</span><span class="sxs-lookup"><span data-stu-id="d3c19-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="d3c19-153">boolean</span><span class="sxs-lookup"><span data-stu-id="d3c19-153">boolean</span></span>  |  <span data-ttu-id="d3c19-154">Non</span><span class="sxs-lookup"><span data-stu-id="d3c19-154">No</span></span><br/><br/><span data-ttu-id="d3c19-155">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="d3c19-155">Default value is 4.</span></span>  |  <span data-ttu-id="d3c19-156">Si `true`, Excel appelle le gestionnaire `onCanceled` à chaque fois que l’utilisateur exécute une action qui a pour effet l’annulation de la fonction ; par exemple, déclencher manuellement le recalcul, ou modifier une cellule référencée par la fonction.</span><span class="sxs-lookup"><span data-stu-id="d3c19-156">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="d3c19-157">Si vous utilisez cette option, Excel appellera la fonction JavaScript avec un paramètre `caller` additionnel.</span><span class="sxs-lookup"><span data-stu-id="d3c19-157">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="d3c19-158">(Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`).</span><span class="sxs-lookup"><span data-stu-id="d3c19-158">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="d3c19-159">Dans le corps de la fonction, un gestionnaire doit être affecté au membre `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="d3c19-159">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="d3c19-160">Pour plus d’informations, voir [Annulation d’une fonction](custom-functions-overview.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="d3c19-160">For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="d3c19-161">boolean</span><span class="sxs-lookup"><span data-stu-id="d3c19-161">boolean</span></span>  |  <span data-ttu-id="d3c19-162">Non</span><span class="sxs-lookup"><span data-stu-id="d3c19-162">No</span></span><br/><br/><span data-ttu-id="d3c19-163">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="d3c19-163">Default value is 4.</span></span>  |  <span data-ttu-id="d3c19-164">Si `true`, la fonction peut générer une sortie plusieurs fois dans la cellule même lorsqu’elle n’est invoquée qu’une seule fois.</span><span class="sxs-lookup"><span data-stu-id="d3c19-164">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="d3c19-165">Cette option est utile pour les sources de données en évolution rapide, telles que le cours d'une action.</span><span class="sxs-lookup"><span data-stu-id="d3c19-165">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="d3c19-166">Si vous utilisez cette option, Excel appellera la fonction JavaScript avec un paramètre `caller` additionnel.</span><span class="sxs-lookup"><span data-stu-id="d3c19-166">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="d3c19-167">(Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`).</span><span class="sxs-lookup"><span data-stu-id="d3c19-167">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="d3c19-168">La fonction ne devrait pas avoir d’instruction `return`.</span><span class="sxs-lookup"><span data-stu-id="d3c19-168">The function should have no `return` statement.</span></span> <span data-ttu-id="d3c19-169">Au lieu de cela, la valeur du résultat est passée comme argument à la méthode de rappel `caller.setResult`.</span><span class="sxs-lookup"><span data-stu-id="d3c19-169">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="d3c19-170">Pour plus d’informations, voir [Fonctions de flux](custom-functions-overview.md#streamed-functions).</span><span class="sxs-lookup"><span data-stu-id="d3c19-170">For more information, see [Excel functions by category](custom-functions-overview.md#streamed-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="d3c19-171">parameters</span><span class="sxs-lookup"><span data-stu-id="d3c19-171">parameters</span></span>

<span data-ttu-id="d3c19-172">La propriété `parameters` est un tableau d’objets parameter.</span><span class="sxs-lookup"><span data-stu-id="d3c19-172">The `parameters` property is an array of objects.</span></span> <span data-ttu-id="d3c19-173">Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="d3c19-173">The following table lists the properties of the SP.FieldRatingScale object.</span></span>

|  <span data-ttu-id="d3c19-174">Propriété</span><span class="sxs-lookup"><span data-stu-id="d3c19-174">Property</span></span>  |  <span data-ttu-id="d3c19-175">Type de données</span><span class="sxs-lookup"><span data-stu-id="d3c19-175">Data type</span></span>  |  <span data-ttu-id="d3c19-176">Requis</span><span class="sxs-lookup"><span data-stu-id="d3c19-176">Required</span></span>  |  <span data-ttu-id="d3c19-177">Description</span><span class="sxs-lookup"><span data-stu-id="d3c19-177">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="d3c19-178">chaîne</span><span class="sxs-lookup"><span data-stu-id="d3c19-178">string</span></span>  |  <span data-ttu-id="d3c19-179">Non</span><span class="sxs-lookup"><span data-stu-id="d3c19-179">No</span></span> |  <span data-ttu-id="d3c19-180">Description du paramètre.</span><span class="sxs-lookup"><span data-stu-id="d3c19-180">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="d3c19-181">chaîne</span><span class="sxs-lookup"><span data-stu-id="d3c19-181">string</span></span>  |  <span data-ttu-id="d3c19-182">Non</span><span class="sxs-lookup"><span data-stu-id="d3c19-182">No</span></span>  |  <span data-ttu-id="d3c19-183">Doit être **scalar** (une valeur non tableau) ou **matrix** (tableau à deux dimensions).</span><span class="sxs-lookup"><span data-stu-id="d3c19-183">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="d3c19-184">chaîne</span><span class="sxs-lookup"><span data-stu-id="d3c19-184">string</span></span>  |  <span data-ttu-id="d3c19-185">Oui</span><span class="sxs-lookup"><span data-stu-id="d3c19-185">Yes</span></span>  |  <span data-ttu-id="d3c19-186">Nom du paramètre.</span><span class="sxs-lookup"><span data-stu-id="d3c19-186">The name of the parameter.</span></span> <span data-ttu-id="d3c19-187">Ce nom est affiché dans l’IntelliSense d’Excel.</span><span class="sxs-lookup"><span data-stu-id="d3c19-187">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="d3c19-188">chaîne</span><span class="sxs-lookup"><span data-stu-id="d3c19-188">string</span></span>  |  <span data-ttu-id="d3c19-189">Non</span><span class="sxs-lookup"><span data-stu-id="d3c19-189">No</span></span>  |  <span data-ttu-id="d3c19-190">Type de données du paramètre.</span><span class="sxs-lookup"><span data-stu-id="d3c19-190">The data type of the parameter.</span></span> <span data-ttu-id="d3c19-191">Doit être **boolean**, **number** ou **string**.</span><span class="sxs-lookup"><span data-stu-id="d3c19-191">Must be "boolean", "number", or "string".</span></span>  |

## <a name="result"></a><span data-ttu-id="d3c19-192">result</span><span class="sxs-lookup"><span data-stu-id="d3c19-192">result</span></span>

<span data-ttu-id="d3c19-193">L’objet `results` définit le type de l’information renvoyée par la fonction.</span><span class="sxs-lookup"><span data-stu-id="d3c19-193">The `results` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="d3c19-194">Le tableau suivant répertorie les propriétés de l’objet `result`.</span><span class="sxs-lookup"><span data-stu-id="d3c19-194">The following table lists the properties of the</span></span>

|  <span data-ttu-id="d3c19-195">Propriété</span><span class="sxs-lookup"><span data-stu-id="d3c19-195">Property</span></span>  |  <span data-ttu-id="d3c19-196">Type de données</span><span class="sxs-lookup"><span data-stu-id="d3c19-196">Data type</span></span>  |  <span data-ttu-id="d3c19-197">Requis</span><span class="sxs-lookup"><span data-stu-id="d3c19-197">Required</span></span>  |  <span data-ttu-id="d3c19-198">Description</span><span class="sxs-lookup"><span data-stu-id="d3c19-198">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="d3c19-199">chaîne</span><span class="sxs-lookup"><span data-stu-id="d3c19-199">string</span></span>  |  <span data-ttu-id="d3c19-200">Non</span><span class="sxs-lookup"><span data-stu-id="d3c19-200">No</span></span>  |  <span data-ttu-id="d3c19-201">Doit être **scalar** (une valeur non tableau) ou **matrix** (tableau à deux dimensions).</span><span class="sxs-lookup"><span data-stu-id="d3c19-201">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="d3c19-202">chaîne</span><span class="sxs-lookup"><span data-stu-id="d3c19-202">string</span></span>  |  <span data-ttu-id="d3c19-203">Oui</span><span class="sxs-lookup"><span data-stu-id="d3c19-203">Yes</span></span>  |  <span data-ttu-id="d3c19-204">Type de données du paramètre.</span><span class="sxs-lookup"><span data-stu-id="d3c19-204">The data type of the parameter.</span></span> <span data-ttu-id="d3c19-205">Doit être **boolean**, **number** ou **string**.</span><span class="sxs-lookup"><span data-stu-id="d3c19-205">Must be "boolean", "number", or "string".</span></span>  |

## <a name="see-also"></a><span data-ttu-id="d3c19-206">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="d3c19-206">See also</span></span>

* [<span data-ttu-id="d3c19-207">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="d3c19-207">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="d3c19-208">Runtime pour les fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="d3c19-208">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="d3c19-209">Meilleures pratiques pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="d3c19-209">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="d3c19-210">Didacticiel sur les fonctions personnalisées d’Excel</span><span class="sxs-lookup"><span data-stu-id="d3c19-210">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)