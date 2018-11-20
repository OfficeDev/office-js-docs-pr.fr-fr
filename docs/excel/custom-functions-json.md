---
ms.date: 10/17/2018
description: Définissez des métadonnées pour des fonctions personnalisées dans Excel.
title: Métadonnées pour des fonctions personnalisées dans Excel
ms.openlocfilehash: 0c77474188a2deefd23a73bb64e87569bb1fa52a
ms.sourcegitcommit: 2ac7d64bb2db75ace516a604866850fce5cb2174
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/14/2018
ms.locfileid: "26298543"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="91ea2-103">Métadonnées de fonctions personnalisées (aperçu)</span><span class="sxs-lookup"><span data-stu-id="91ea2-103">Custom functions metadata (preview)</span></span>

<span data-ttu-id="91ea2-104">Lorsque vous définissez des [fonctions personnalisées](custom-functions-overview.md) dans votre complément Excel, votre projet de complément doit inclure un fichier de métadonnées JSON qui fournit les informations dont Excel a besoin pour enregistrer les fonctions personnalisées et les rendre disponibles aux utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="91ea2-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="91ea2-105">Cet article décrit le format du fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="91ea2-105">This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="91ea2-106">Pour plus d’informations sur les autres fichiers à inclure dans votre projet de complément afin d’activer les fonctions personnalisées, voir [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="91ea2-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="91ea2-107">Exemple de métadonnées</span><span class="sxs-lookup"><span data-stu-id="91ea2-107">Example metadata</span></span>

<span data-ttu-id="91ea2-108">L’exemple suivant montre le contenu d’un fichier de métadonnées JSON pour un complément qui définit des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="91ea2-108">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="91ea2-109">Les sections qui suivent cet exemple fournissent des informations détaillées sur les propriétés individuelles au sein de cet exemple JSON.</span><span class="sxs-lookup"><span data-stu-id="91ea2-109">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="91ea2-110">Un exemple de fichier JSON complet est disponible dans le dépôt GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="91ea2-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="91ea2-111">fonctions</span><span class="sxs-lookup"><span data-stu-id="91ea2-111">functions</span></span> 

<span data-ttu-id="91ea2-112">La propriété `functions` est un tableau d’objets de fonction personnalisés.</span><span class="sxs-lookup"><span data-stu-id="91ea2-112">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="91ea2-113">Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="91ea2-113">The following table lists the properties of the SP.FieldRatingScale object.</span></span>

|  <span data-ttu-id="91ea2-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="91ea2-114">Property</span></span>  |  <span data-ttu-id="91ea2-115">Type de données</span><span class="sxs-lookup"><span data-stu-id="91ea2-115">Data type</span></span>  |  <span data-ttu-id="91ea2-116">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="91ea2-116">Required</span></span>  |  <span data-ttu-id="91ea2-117">Description</span><span class="sxs-lookup"><span data-stu-id="91ea2-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="91ea2-118">string</span><span class="sxs-lookup"><span data-stu-id="91ea2-118">string</span></span>  |  <span data-ttu-id="91ea2-119">Non</span><span class="sxs-lookup"><span data-stu-id="91ea2-119">No</span></span>  |  <span data-ttu-id="91ea2-120">Description de la fonction que voient les utilisateurs finaux dans Excel.</span><span class="sxs-lookup"><span data-stu-id="91ea2-120">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="91ea2-121">Par exemple, **convertit une valeur Celsius en valeur Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="91ea2-121">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="91ea2-122">string</span><span class="sxs-lookup"><span data-stu-id="91ea2-122">string</span></span>  |   <span data-ttu-id="91ea2-123">Non</span><span class="sxs-lookup"><span data-stu-id="91ea2-123">No</span></span>  |  <span data-ttu-id="91ea2-124">URL fournissant des informations sur la fonction</span><span class="sxs-lookup"><span data-stu-id="91ea2-124">URL that provides information about the function.</span></span> <span data-ttu-id="91ea2-125">(elle est affichée dans un volet des tâches). Par exemple, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="91ea2-125">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="91ea2-126">string</span><span class="sxs-lookup"><span data-stu-id="91ea2-126">string</span></span> | <span data-ttu-id="91ea2-127">Oui</span><span class="sxs-lookup"><span data-stu-id="91ea2-127">Yes</span></span> | <span data-ttu-id="91ea2-128">Un ID unique pour la fonction.</span><span class="sxs-lookup"><span data-stu-id="91ea2-128">A unique ID for the group.</span></span> <span data-ttu-id="91ea2-129">Cet ID peut contenir uniquement des points et caractères alphanumériques et ne doit pas être modifié une fois défini.</span><span class="sxs-lookup"><span data-stu-id="91ea2-129">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="91ea2-130">string</span><span class="sxs-lookup"><span data-stu-id="91ea2-130">string</span></span>  |  <span data-ttu-id="91ea2-131">Oui</span><span class="sxs-lookup"><span data-stu-id="91ea2-131">Yes</span></span>  |  <span data-ttu-id="91ea2-132">Nom de la fonction que voient les utilisateurs finaux dans Excel.</span><span class="sxs-lookup"><span data-stu-id="91ea2-132">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="91ea2-133">Dans Excel, le nom de la fonction sera précédé de l’espace de noms de fonctions personnalisées spécifié dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="91ea2-133">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="91ea2-134">object</span><span class="sxs-lookup"><span data-stu-id="91ea2-134">object</span></span>  |  <span data-ttu-id="91ea2-135">Non</span><span class="sxs-lookup"><span data-stu-id="91ea2-135">No</span></span>  |  <span data-ttu-id="91ea2-136">Vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction.</span><span class="sxs-lookup"><span data-stu-id="91ea2-136">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="91ea2-137">Pour plus d’informations, voir [objet options](#options-object).</span><span class="sxs-lookup"><span data-stu-id="91ea2-137">See object load [options](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="91ea2-138">array</span><span class="sxs-lookup"><span data-stu-id="91ea2-138">array</span></span>  |  <span data-ttu-id="91ea2-139">Oui</span><span class="sxs-lookup"><span data-stu-id="91ea2-139">Yes</span></span>  |  <span data-ttu-id="91ea2-140">Tableau qui définit les paramètres d’entrée de la fonction.</span><span class="sxs-lookup"><span data-stu-id="91ea2-140">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="91ea2-141">Pour plus d’informations, voir [tableau de paramètres](#parameters-array).</span><span class="sxs-lookup"><span data-stu-id="91ea2-141">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="91ea2-142">object</span><span class="sxs-lookup"><span data-stu-id="91ea2-142">object</span></span>  |  <span data-ttu-id="91ea2-143">Oui</span><span class="sxs-lookup"><span data-stu-id="91ea2-143">Yes</span></span>  |  <span data-ttu-id="91ea2-144">Objet qui définit le type d’informations renvoyées par la fonction.</span><span class="sxs-lookup"><span data-stu-id="91ea2-144">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="91ea2-145">Pour plus d’informations, voir [objet résultat](#result-object).</span><span class="sxs-lookup"><span data-stu-id="91ea2-145">See object load [options](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="91ea2-146">options</span><span class="sxs-lookup"><span data-stu-id="91ea2-146">options</span></span>

<span data-ttu-id="91ea2-147">L’objet `options` vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction.</span><span class="sxs-lookup"><span data-stu-id="91ea2-147">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="91ea2-148">Le tableau suivant répertorie les propriétés de l’objet `options`.</span><span class="sxs-lookup"><span data-stu-id="91ea2-148">The following table lists the properties of the</span></span>

|  <span data-ttu-id="91ea2-149">Propriété</span><span class="sxs-lookup"><span data-stu-id="91ea2-149">Property</span></span>  |  <span data-ttu-id="91ea2-150">Type de données</span><span class="sxs-lookup"><span data-stu-id="91ea2-150">Data type</span></span>  |  <span data-ttu-id="91ea2-151">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="91ea2-151">Required</span></span>  |  <span data-ttu-id="91ea2-152">Description</span><span class="sxs-lookup"><span data-stu-id="91ea2-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="91ea2-153">boolean</span><span class="sxs-lookup"><span data-stu-id="91ea2-153">boolean</span></span>  |  <span data-ttu-id="91ea2-154">Non</span><span class="sxs-lookup"><span data-stu-id="91ea2-154">No</span></span><br/><br/><span data-ttu-id="91ea2-155">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="91ea2-155">Default value is `false`.</span></span>  |  <span data-ttu-id="91ea2-156">Si la valeur est `true`, Excel appelle le gestionnaire `onCanceled` chaque fois que l’utilisateur effectue une action ayant pour effet d’annuler la fonction, par exemple, en déclenchant manuellement un recalcul ou en modifiant une cellule référencée par la fonction.</span><span class="sxs-lookup"><span data-stu-id="91ea2-156">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="91ea2-157">Si vous utilisez cette option, Excel appelle la fonction JavaScript avec un paramètre `caller` supplémentaire</span><span class="sxs-lookup"><span data-stu-id="91ea2-157">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="91ea2-158">(n’enregistrez ***pas*** ce paramètre dans la propriété `parameters`).</span><span class="sxs-lookup"><span data-stu-id="91ea2-158">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="91ea2-159">Dans le corps de la fonction, un gestionnaire doit être attribué au membre `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="91ea2-159">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="91ea2-160">Pour plus d’informations, voir [Annuler une fonction](custom-functions-overview.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="91ea2-160">For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="91ea2-161">boolean</span><span class="sxs-lookup"><span data-stu-id="91ea2-161">boolean</span></span>  |  <span data-ttu-id="91ea2-162">Non</span><span class="sxs-lookup"><span data-stu-id="91ea2-162">No</span></span><br/><br/><span data-ttu-id="91ea2-163">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="91ea2-163">Default value is `false`.</span></span>  |  <span data-ttu-id="91ea2-164">Si la valeur est `true`, la fonction peut envoyer une sortie à la cellule à plusieurs reprises, même en cas d’appel unique.</span><span class="sxs-lookup"><span data-stu-id="91ea2-164">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="91ea2-165">Cette option est utile pour des sources de données qui changent rapidement, telles que des valeurs boursières.</span><span class="sxs-lookup"><span data-stu-id="91ea2-165">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="91ea2-166">Si vous utilisez cette option, Excel appelle la fonction JavaScript avec un paramètre `caller` supplémentaire</span><span class="sxs-lookup"><span data-stu-id="91ea2-166">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="91ea2-167">(n’enregistrez ***pas*** ce paramètre dans la propriété `parameters`).</span><span class="sxs-lookup"><span data-stu-id="91ea2-167">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="91ea2-168">La fonction ne doit pas utiliser d’instruction `return`.</span><span class="sxs-lookup"><span data-stu-id="91ea2-168">The function should have no `return` statement.</span></span> <span data-ttu-id="91ea2-169">Au lieu de cela, la valeur obtenue est transmise en tant qu’argument de la méthode de rappel `caller.setResult`.</span><span class="sxs-lookup"><span data-stu-id="91ea2-169">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="91ea2-170">Pour plus d’informations, voir [Diffusion en continu de fonctions](custom-functions-overview.md#streaming-functions).</span><span class="sxs-lookup"><span data-stu-id="91ea2-170">For more information, see [Streaming functions](custom-functions-overview.md#streaming-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="91ea2-171">parameters</span><span class="sxs-lookup"><span data-stu-id="91ea2-171">parameters</span></span>

<span data-ttu-id="91ea2-172">La propriété `parameters` est un tableau d’objets paramètre.</span><span class="sxs-lookup"><span data-stu-id="91ea2-172">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="91ea2-173">Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="91ea2-173">The following table lists the properties of the SP.FieldRatingScale object.</span></span>

|  <span data-ttu-id="91ea2-174">Propriété</span><span class="sxs-lookup"><span data-stu-id="91ea2-174">Property</span></span>  |  <span data-ttu-id="91ea2-175">Type de données</span><span class="sxs-lookup"><span data-stu-id="91ea2-175">Data type</span></span>  |  <span data-ttu-id="91ea2-176">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="91ea2-176">Required</span></span>  |  <span data-ttu-id="91ea2-177">Description</span><span class="sxs-lookup"><span data-stu-id="91ea2-177">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="91ea2-178">string</span><span class="sxs-lookup"><span data-stu-id="91ea2-178">string</span></span>  |  <span data-ttu-id="91ea2-179">Non</span><span class="sxs-lookup"><span data-stu-id="91ea2-179">No</span></span> |  <span data-ttu-id="91ea2-180">Description du paramètre.</span><span class="sxs-lookup"><span data-stu-id="91ea2-180">A description of the value.</span></span> <span data-ttu-id="91ea2-181">S’affiche dans intelliSense d’Excel.</span><span class="sxs-lookup"><span data-stu-id="91ea2-181">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="91ea2-182">string</span><span class="sxs-lookup"><span data-stu-id="91ea2-182">string</span></span>  |  <span data-ttu-id="91ea2-183">Non</span><span class="sxs-lookup"><span data-stu-id="91ea2-183">No</span></span>  |  <span data-ttu-id="91ea2-184">Doit être **scalaire** (valeur autre que de tableau) ou **matrice** (tableau bidimensionnel).</span><span class="sxs-lookup"><span data-stu-id="91ea2-184">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="91ea2-185">string</span><span class="sxs-lookup"><span data-stu-id="91ea2-185">string</span></span>  |  <span data-ttu-id="91ea2-186">Oui</span><span class="sxs-lookup"><span data-stu-id="91ea2-186">Yes</span></span>  |  <span data-ttu-id="91ea2-187">Le nom du paramètre.</span><span class="sxs-lookup"><span data-stu-id="91ea2-187">The name of the parameter.</span></span> <span data-ttu-id="91ea2-188">Ce nom s’affiche dans intelliSense d’Excel.</span><span class="sxs-lookup"><span data-stu-id="91ea2-188">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="91ea2-189">string</span><span class="sxs-lookup"><span data-stu-id="91ea2-189">string</span></span>  |  <span data-ttu-id="91ea2-190">Non</span><span class="sxs-lookup"><span data-stu-id="91ea2-190">No</span></span>  |  <span data-ttu-id="91ea2-191">Type de données du paramètre.</span><span class="sxs-lookup"><span data-stu-id="91ea2-191">The System data type of the parameter.</span></span> <span data-ttu-id="91ea2-192">Peut être **boolean**, **number**, **string** ou **any** qui vous permet d’utiliser n’importe lequel des trois types précédents.</span><span class="sxs-lookup"><span data-stu-id="91ea2-192">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="91ea2-193">Si cette propriété n’est pas spécifiée, le type de données par défaut est **any**.</span><span class="sxs-lookup"><span data-stu-id="91ea2-193">If this property is not specified, the data type defaults to **any**.</span></span> |

## <a name="result"></a><span data-ttu-id="91ea2-194">result</span><span class="sxs-lookup"><span data-stu-id="91ea2-194">result</span></span>

<span data-ttu-id="91ea2-195">L’objet `result` définit le type des informations renvoyées par la fonction.</span><span class="sxs-lookup"><span data-stu-id="91ea2-195">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="91ea2-196">Le tableau suivant répertorie les propriétés de l’objet `result`.</span><span class="sxs-lookup"><span data-stu-id="91ea2-196">The following table lists the properties of the</span></span>

|  <span data-ttu-id="91ea2-197">Propriété</span><span class="sxs-lookup"><span data-stu-id="91ea2-197">Property</span></span>  |  <span data-ttu-id="91ea2-198">Type de données</span><span class="sxs-lookup"><span data-stu-id="91ea2-198">Data type</span></span>  |  <span data-ttu-id="91ea2-199">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="91ea2-199">Required</span></span>  |  <span data-ttu-id="91ea2-200">Description</span><span class="sxs-lookup"><span data-stu-id="91ea2-200">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="91ea2-201">string</span><span class="sxs-lookup"><span data-stu-id="91ea2-201">string</span></span>  |  <span data-ttu-id="91ea2-202">Non</span><span class="sxs-lookup"><span data-stu-id="91ea2-202">No</span></span>  |  <span data-ttu-id="91ea2-203">Doit être **scalaire** (valeur autre que de tableau) ou **matrice** (tableau bidimensionnel).</span><span class="sxs-lookup"><span data-stu-id="91ea2-203">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="91ea2-204">string</span><span class="sxs-lookup"><span data-stu-id="91ea2-204">string</span></span>  |  <span data-ttu-id="91ea2-205">Oui</span><span class="sxs-lookup"><span data-stu-id="91ea2-205">Yes</span></span>  |  <span data-ttu-id="91ea2-206">Type de données du paramètre.</span><span class="sxs-lookup"><span data-stu-id="91ea2-206">The System data type of the parameter.</span></span> <span data-ttu-id="91ea2-207">Doit être **boolean**, **number**, **string** ou **any** qui vous permet d’utiliser n’importe lequel des trois types précédents.</span><span class="sxs-lookup"><span data-stu-id="91ea2-207">Must be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> |

## <a name="see-also"></a><span data-ttu-id="91ea2-208">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="91ea2-208">See also</span></span>

* [<span data-ttu-id="91ea2-209">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="91ea2-209">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="91ea2-210">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="91ea2-210">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="91ea2-211">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="91ea2-211">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="91ea2-212">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="91ea2-212">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
