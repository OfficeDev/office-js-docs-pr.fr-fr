---
ms.date: 09/20/2018
description: Définir les métadonnées pour des fonctions personnalisées dans Excel.
title: Métadonnées pour des fonctions personnalisées dans Excel
ms.openlocfilehash: 815b0c6e65966867d9e5d953a40ffc705a63ee63
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/21/2018
ms.locfileid: "24062143"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="c2696-103">Métadonnées des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="c2696-103">Custom functions metadata</span></span>

<span data-ttu-id="c2696-104">Lorsque vous définissez des[fonctions personnalisées](custom-functions-overview.md) dans votre complément Excel, votre projet de complément doit inclure un fichier de métadonnées JSON qui fournit les informations nécessaires pour inscrire les fonctions personnalisées et de les rendre disponibles pour les utilisateurs finaux dans Excel.</span><span class="sxs-lookup"><span data-stu-id="c2696-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end-users.</span></span> <span data-ttu-id="c2696-105">Cet article décrit le format du fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="c2696-105">This article describes the format of the JSON file with examples.</span></span>

> [!NOTE]
> <span data-ttu-id="c2696-106">Pour plus d’informations sur les autres fichiers que vous devez inclure dans votre projet de complément pour activer les fonctions personnalisées, voir [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md#learn-the-basics).</span><span class="sxs-lookup"><span data-stu-id="c2696-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md#learn-the-basics).</span></span>

## <a name="example-metadata"></a><span data-ttu-id="c2696-107">Métadonnées d’exemple</span><span class="sxs-lookup"><span data-stu-id="c2696-107">Example metadata</span></span>

<span data-ttu-id="c2696-108">L’exemple suivant montre le contenu d’un fichier de métadonnées JSON pour un complément qui définit des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="c2696-108">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="c2696-109">Les sections qui suivent cet exemple fournissent des informations détaillées sur les propriétés individuelles dans cet exemple JSON.</span><span class="sxs-lookup"><span data-stu-id="c2696-109">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

```json
{
    "functions": [
        {
            "id": "ADD42",
            "name": "ADD42",
            "description":  "Adds 42 to the input number",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "ADD42ASYNC",
            "name": "ADD42ASYNC",
            "description":  "asynchronously wait 250ms, then add 42",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "ISEVEN",
            "name": "ISEVEN", 
            "description":  "Determines whether a number is even",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "boolean",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "the number to be evaluated",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "GETDAY",
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": []
        },
        {
            "id": "INCREMENTVALUE",
            "name": "INCREMENTVALUE", 
            "description":  "Counts up from zero",
            "helpUrl": "http://dev.office.com",
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
            "description":  "gets the second highest number from a range",
            "helpUrl": "http://dev.office.com",
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
> <span data-ttu-id="c2696-110">Un fichier d’exemple JSON complet est disponible dans le [référentiel GitHub OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="c2696-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions GitHub repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions"></a><span data-ttu-id="c2696-111">functions</span><span class="sxs-lookup"><span data-stu-id="c2696-111">functions</span></span> 

<span data-ttu-id="c2696-112">La propriété `functions` est un tableau d’objets de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="c2696-112">The `functions` property is an array of objects.</span></span> <span data-ttu-id="c2696-113">Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="c2696-113">The following table lists the properties of the SP.FieldRatingScale object.</span></span>

|  <span data-ttu-id="c2696-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="c2696-114">Property</span></span>  |  <span data-ttu-id="c2696-115">Type de données</span><span class="sxs-lookup"><span data-stu-id="c2696-115">Data type</span></span>  |  <span data-ttu-id="c2696-116">Requis</span><span class="sxs-lookup"><span data-stu-id="c2696-116">Required</span></span>  |  <span data-ttu-id="c2696-117">Description</span><span class="sxs-lookup"><span data-stu-id="c2696-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="c2696-118">string</span><span class="sxs-lookup"><span data-stu-id="c2696-118">string</span></span>  |  <span data-ttu-id="c2696-119">Non</span><span class="sxs-lookup"><span data-stu-id="c2696-119">No</span></span>  |  <span data-ttu-id="c2696-120">Une description de la fonction apparaissant dans l’interface utilisateur Excel.</span><span class="sxs-lookup"><span data-stu-id="c2696-120">A description of the function that appears in the Excel UI.</span></span> <span data-ttu-id="c2696-121">Par exemple, **Convertit une valeur Celsius en Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="c2696-121">For example, "Converts a Celsius value to Fahrenheit".</span></span> |
|  `helpUrl`  |  <span data-ttu-id="c2696-122">string</span><span class="sxs-lookup"><span data-stu-id="c2696-122">string</span></span>  |   <span data-ttu-id="c2696-123">Non</span><span class="sxs-lookup"><span data-stu-id="c2696-123">No</span></span>  |  <span data-ttu-id="c2696-124">L’URL où vos utilisateurs peuvent obtenir de l’aide sur la fonction.</span><span class="sxs-lookup"><span data-stu-id="c2696-124">URL where your users can get help about the function.</span></span> <span data-ttu-id="c2696-125">(Elle est affichée dans un volet Office.) Par exemple, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="c2696-125">(It is displayed in a taskpane.) For example, "http://contoso.com/help/convertcelsiustofahrenheit.html"</span></span> |
| `id`     | <span data-ttu-id="c2696-126">string</span><span class="sxs-lookup"><span data-stu-id="c2696-126">string</span></span> | <span data-ttu-id="c2696-127">Oui</span><span class="sxs-lookup"><span data-stu-id="c2696-127">Yes</span></span> | <span data-ttu-id="c2696-128">ID unique de la fonction.</span><span class="sxs-lookup"><span data-stu-id="c2696-128">A unique ID for the group.</span></span> <span data-ttu-id="c2696-129">Cet ID ne doit pas être modifié après sa définition.</span><span class="sxs-lookup"><span data-stu-id="c2696-129">This ID should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="c2696-130">string</span><span class="sxs-lookup"><span data-stu-id="c2696-130">string</span></span>  |  <span data-ttu-id="c2696-131">Oui</span><span class="sxs-lookup"><span data-stu-id="c2696-131">Yes</span></span>  |  <span data-ttu-id="c2696-132">Le nom de la fonction telle qu'elle apparaîtra (préfixée d'un espace de nom) dans l'interface utilisateur Excel lorsqu'un utilisateur sélectionne une fonction.</span><span class="sxs-lookup"><span data-stu-id="c2696-132">The name of the function as it will appear (prepended with a namespace) in the Excel UI when a user is selecting a function.</span></span> <span data-ttu-id="c2696-133">Il n’a pas besoin d’être le même que le nom de la fonction telle que définie dans le JavaScript.</span><span class="sxs-lookup"><span data-stu-id="c2696-133">It should be the same as the function's name where it is defined in the JavaScript.</span></span> |
|  `options`  |  <span data-ttu-id="c2696-134">object</span><span class="sxs-lookup"><span data-stu-id="c2696-134">object</span></span>  |  <span data-ttu-id="c2696-135">Non</span><span class="sxs-lookup"><span data-stu-id="c2696-135">No</span></span>  |  <span data-ttu-id="c2696-136">Vous permet de personnaliser certains aspects de la façon dont Excel exécute la fonction, et quand.</span><span class="sxs-lookup"><span data-stu-id="c2696-136">The  property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="c2696-137">Voir [objet options](#options-object) pour plus de détails.</span><span class="sxs-lookup"><span data-stu-id="c2696-137">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="c2696-138">array</span><span class="sxs-lookup"><span data-stu-id="c2696-138">array</span></span>  |  <span data-ttu-id="c2696-139">Oui</span><span class="sxs-lookup"><span data-stu-id="c2696-139">Yes</span></span>  |  <span data-ttu-id="c2696-140">Tableau qui définit les paramètres d’entrée de la fonction.</span><span class="sxs-lookup"><span data-stu-id="c2696-140">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="c2696-141">Voir[tableau parameters](#parameters-array) pour plus de détails.</span><span class="sxs-lookup"><span data-stu-id="c2696-141">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="c2696-142">object</span><span class="sxs-lookup"><span data-stu-id="c2696-142">object</span></span>  |  <span data-ttu-id="c2696-143">Oui</span><span class="sxs-lookup"><span data-stu-id="c2696-143">Yes</span></span>  |  <span data-ttu-id="c2696-144">Objet qui définit le type de l’information renvoyée par la fonction.</span><span class="sxs-lookup"><span data-stu-id="c2696-144">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="c2696-145">Voir [objet result](#result-object) pour plus de détails.</span><span class="sxs-lookup"><span data-stu-id="c2696-145">See [result object](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="c2696-146">options</span><span class="sxs-lookup"><span data-stu-id="c2696-146">options</span></span>

<span data-ttu-id="c2696-147">L’objet `options` vous permet de personnaliser certains aspects de la façon dont Excel exécute la fonction, et quand.</span><span class="sxs-lookup"><span data-stu-id="c2696-147">The `options` property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="c2696-148">Le tableau suivant répertorie les propriétés de l’objet `options`.</span><span class="sxs-lookup"><span data-stu-id="c2696-148">The following table lists the properties of the</span></span>

|  <span data-ttu-id="c2696-149">Propriété</span><span class="sxs-lookup"><span data-stu-id="c2696-149">Property</span></span>  |  <span data-ttu-id="c2696-150">Type de données</span><span class="sxs-lookup"><span data-stu-id="c2696-150">Data type</span></span>  |  <span data-ttu-id="c2696-151">Requis</span><span class="sxs-lookup"><span data-stu-id="c2696-151">Required</span></span>  |  <span data-ttu-id="c2696-152">Description</span><span class="sxs-lookup"><span data-stu-id="c2696-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="c2696-153">boolean</span><span class="sxs-lookup"><span data-stu-id="c2696-153">boolean</span></span>  |  <span data-ttu-id="c2696-154">Non, la valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="c2696-154">No, default is `false`.</span></span>  |  <span data-ttu-id="c2696-155">Lorsqu’`true`Excel appelle le `onCanceled` gestionnaire au moment où l'utilisateur prend une action visant par exemple à annuler la fonction, le déclenchement manuel du recalcul ou la modification d’une cellule est référencée par cette fonction.</span><span class="sxs-lookup"><span data-stu-id="c2696-155">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="c2696-156">Si vous utilisez cette option, Excel appellera la fonction JavaScript avec un paramètre `caller` additionnel.</span><span class="sxs-lookup"><span data-stu-id="c2696-156">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="c2696-157">(Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`).</span><span class="sxs-lookup"><span data-stu-id="c2696-157">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="c2696-158">Dans le corps de la fonction, un gestionnaire doit être affecté au membre `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="c2696-158">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="c2696-159">Pour plus d’informations, voir [Annulation d’une fonction](custom-functions-overview.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="c2696-159">For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="c2696-160">boolean</span><span class="sxs-lookup"><span data-stu-id="c2696-160">boolean</span></span>  |  <span data-ttu-id="c2696-161">Non, la valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="c2696-161">No, default is `false`.</span></span>  |  <span data-ttu-id="c2696-162">Si `true`, la fonction peut générer une sortie plusieurs fois dans la cellule même lorsqu'elle n'est invoquée qu'une seule fois.</span><span class="sxs-lookup"><span data-stu-id="c2696-162">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="c2696-163">Cette option est utile pour les sources de données en évolution rapide, telles que le cours d'une action.</span><span class="sxs-lookup"><span data-stu-id="c2696-163">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="c2696-164">Si vous utilisez cette option, Excel appellera la fonction JavaScript avec un paramètre `caller` additionnel.</span><span class="sxs-lookup"><span data-stu-id="c2696-164">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="c2696-165">(Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`).</span><span class="sxs-lookup"><span data-stu-id="c2696-165">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="c2696-166">La fonction ne devrait pas avoir de `return` déclaration.</span><span class="sxs-lookup"><span data-stu-id="c2696-166">The function should have no `return` statement.</span></span> <span data-ttu-id="c2696-167">Au lieu de cela, la valeur du résultat est passée comme argument à la méthode de rappel `caller.setResult`.</span><span class="sxs-lookup"><span data-stu-id="c2696-167">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="c2696-168">Pour plus d’informations, voir [Fonctions de flux](custom-functions-overview.md#streamed-functions).</span><span class="sxs-lookup"><span data-stu-id="c2696-168">For more information, see [Excel functions by category](custom-functions-overview.md#streamed-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="c2696-169">parameters</span><span class="sxs-lookup"><span data-stu-id="c2696-169">parameters</span></span>

<span data-ttu-id="c2696-170">La propriété `parameters` est un tableau d’objets parameter.</span><span class="sxs-lookup"><span data-stu-id="c2696-170">The `parameters` property is an array of objects.</span></span> <span data-ttu-id="c2696-171">Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="c2696-171">The following table lists the properties of the SP.FieldRatingScale object.</span></span>

|  <span data-ttu-id="c2696-172">Propriété</span><span class="sxs-lookup"><span data-stu-id="c2696-172">Property</span></span>  |  <span data-ttu-id="c2696-173">Type de données</span><span class="sxs-lookup"><span data-stu-id="c2696-173">Data type</span></span>  |  <span data-ttu-id="c2696-174">Requis</span><span class="sxs-lookup"><span data-stu-id="c2696-174">Required</span></span>  |  <span data-ttu-id="c2696-175">Description</span><span class="sxs-lookup"><span data-stu-id="c2696-175">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="c2696-176">string</span><span class="sxs-lookup"><span data-stu-id="c2696-176">string</span></span>  |  <span data-ttu-id="c2696-177">Non</span><span class="sxs-lookup"><span data-stu-id="c2696-177">No</span></span> |  <span data-ttu-id="c2696-178">Une description du paramètre.</span><span class="sxs-lookup"><span data-stu-id="c2696-178">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="c2696-179">string</span><span class="sxs-lookup"><span data-stu-id="c2696-179">string</span></span>  |  <span data-ttu-id="c2696-180">Non</span><span class="sxs-lookup"><span data-stu-id="c2696-180">No</span></span>  |  <span data-ttu-id="c2696-181">Doit être **scalar** (une valeur non tableau) ou **matrix** (tableau à deux dimensions).</span><span class="sxs-lookup"><span data-stu-id="c2696-181">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="c2696-182">string</span><span class="sxs-lookup"><span data-stu-id="c2696-182">string</span></span>  |  <span data-ttu-id="c2696-183">Oui</span><span class="sxs-lookup"><span data-stu-id="c2696-183">Yes</span></span>  |  <span data-ttu-id="c2696-184">Nom du paramètre.</span><span class="sxs-lookup"><span data-stu-id="c2696-184">The name of the parameter.</span></span> <span data-ttu-id="c2696-185">Ce nom est affiché dans l’IntelliSense d’Excel.</span><span class="sxs-lookup"><span data-stu-id="c2696-185">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="c2696-186">string</span><span class="sxs-lookup"><span data-stu-id="c2696-186">string</span></span>  |  <span data-ttu-id="c2696-187">Non</span><span class="sxs-lookup"><span data-stu-id="c2696-187">No</span></span>  |  <span data-ttu-id="c2696-188">Le type de données du paramètre.</span><span class="sxs-lookup"><span data-stu-id="c2696-188">The data type of the parameter.</span></span> <span data-ttu-id="c2696-189">Doit être **boolean**, **number** ou **string**.</span><span class="sxs-lookup"><span data-stu-id="c2696-189">Must be "boolean", "number", or "string".</span></span>  |

## <a name="result"></a><span data-ttu-id="c2696-190">result</span><span class="sxs-lookup"><span data-stu-id="c2696-190">result</span></span>

<span data-ttu-id="c2696-191">L’objet `results` définit le type de l’information renvoyée par la fonction.</span><span class="sxs-lookup"><span data-stu-id="c2696-191">The `results` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="c2696-192">Le tableau suivant répertorie les propriétés de l’objet `result`.</span><span class="sxs-lookup"><span data-stu-id="c2696-192">The following table lists the properties of the</span></span>

|  <span data-ttu-id="c2696-193">Propriété</span><span class="sxs-lookup"><span data-stu-id="c2696-193">Property</span></span>  |  <span data-ttu-id="c2696-194">Type de données</span><span class="sxs-lookup"><span data-stu-id="c2696-194">Data type</span></span>  |  <span data-ttu-id="c2696-195">Requis</span><span class="sxs-lookup"><span data-stu-id="c2696-195">Required</span></span>  |  <span data-ttu-id="c2696-196">Description</span><span class="sxs-lookup"><span data-stu-id="c2696-196">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="c2696-197">string</span><span class="sxs-lookup"><span data-stu-id="c2696-197">string</span></span>  |  <span data-ttu-id="c2696-198">Non</span><span class="sxs-lookup"><span data-stu-id="c2696-198">No</span></span>  |  <span data-ttu-id="c2696-199">Doit être **scalar** (une valeur non tableau) ou **matrix** (tableau à deux dimensions).</span><span class="sxs-lookup"><span data-stu-id="c2696-199">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="c2696-200">string</span><span class="sxs-lookup"><span data-stu-id="c2696-200">string</span></span>  |  <span data-ttu-id="c2696-201">Oui</span><span class="sxs-lookup"><span data-stu-id="c2696-201">Yes</span></span>  |  <span data-ttu-id="c2696-202">Le type de données du paramètre.</span><span class="sxs-lookup"><span data-stu-id="c2696-202">The data type of the parameter.</span></span> <span data-ttu-id="c2696-203">Doit être **boolean**, **number** ou **string**.</span><span class="sxs-lookup"><span data-stu-id="c2696-203">Must be "boolean", "number", or "string".</span></span>  |

## <a name="see-also"></a><span data-ttu-id="c2696-204">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c2696-204">See also</span></span>

* [<span data-ttu-id="c2696-205">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="c2696-205">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="c2696-206">Runtime pour les fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="c2696-206">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="c2696-207">Meilleures pratiques pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="c2696-207">Custom functions best practices</span></span>](custom-functions-best-practices.md)