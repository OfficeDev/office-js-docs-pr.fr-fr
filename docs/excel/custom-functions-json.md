---
ms.date: 11/26/2018
description: Définissez des métadonnées pour des fonctions personnalisées dans Excel.
title: Métadonnées pour des fonctions personnalisées dans Excel
ms.openlocfilehash: 4bdf27173c5e912aa3eba3c8661ba45dd8b453cb
ms.sourcegitcommit: 3007bf57515b0811ff98a7e1518ecc6fc9462276
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/04/2019
ms.locfileid: "27724857"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="7247a-103">Métadonnées de fonctions personnalisées (aperçu)</span><span class="sxs-lookup"><span data-stu-id="7247a-103">Custom functions metadata (preview)</span></span>

<span data-ttu-id="7247a-104">Lorsque vous définissez des [fonctions personnalisées](custom-functions-overview.md) dans votre complément Excel, votre projet de complément doit inclure un fichier de métadonnées JSON qui fournit les informations dont Excel a besoin pour enregistrer les fonctions personnalisées et les rendre disponibles aux utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="7247a-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="7247a-105">Cet article décrit le format du fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="7247a-105">This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="7247a-106">Pour plus d’informations sur les autres fichiers à inclure dans votre projet de complément afin d’activer les fonctions personnalisées, voir [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="7247a-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="7247a-107">Exemple de métadonnées</span><span class="sxs-lookup"><span data-stu-id="7247a-107">Example metadata</span></span>

<span data-ttu-id="7247a-108">L’exemple suivant montre le contenu d’un fichier de métadonnées JSON pour un complément qui définit des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="7247a-108">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="7247a-109">Les sections qui suivent cet exemple fournissent des informations détaillées sur les propriétés individuelles au sein de cet exemple JSON.</span><span class="sxs-lookup"><span data-stu-id="7247a-109">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="7247a-110">Un exemple de fichier JSON complet est disponible dans le dépôt GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="7247a-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="7247a-111">fonctions</span><span class="sxs-lookup"><span data-stu-id="7247a-111">functions</span></span> 

<span data-ttu-id="7247a-112">La propriété `functions` est un tableau d’objets de fonction personnalisés.</span><span class="sxs-lookup"><span data-stu-id="7247a-112">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="7247a-113">Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="7247a-113">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="7247a-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="7247a-114">Property</span></span>  |  <span data-ttu-id="7247a-115">Type de données</span><span class="sxs-lookup"><span data-stu-id="7247a-115">Data type</span></span>  |  <span data-ttu-id="7247a-116">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="7247a-116">Required</span></span>  |  <span data-ttu-id="7247a-117">Description</span><span class="sxs-lookup"><span data-stu-id="7247a-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="7247a-118">string</span><span class="sxs-lookup"><span data-stu-id="7247a-118">string</span></span>  |  <span data-ttu-id="7247a-119">Non</span><span class="sxs-lookup"><span data-stu-id="7247a-119">No</span></span>  |  <span data-ttu-id="7247a-120">Description de la fonction que voient les utilisateurs finaux dans Excel.</span><span class="sxs-lookup"><span data-stu-id="7247a-120">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="7247a-121">Par exemple, **convertit une valeur Celsius en valeur Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="7247a-121">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="7247a-122">string</span><span class="sxs-lookup"><span data-stu-id="7247a-122">string</span></span>  |   <span data-ttu-id="7247a-123">Non</span><span class="sxs-lookup"><span data-stu-id="7247a-123">No</span></span>  |  <span data-ttu-id="7247a-124">URL fournissant des informations sur la fonction</span><span class="sxs-lookup"><span data-stu-id="7247a-124">URL that provides information about the function.</span></span> <span data-ttu-id="7247a-125">(elle est affichée dans un volet des tâches). Par exemple, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="7247a-125">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="7247a-126">string</span><span class="sxs-lookup"><span data-stu-id="7247a-126">string</span></span> | <span data-ttu-id="7247a-127">Oui</span><span class="sxs-lookup"><span data-stu-id="7247a-127">Yes</span></span> | <span data-ttu-id="7247a-128">Un ID unique pour la fonction.</span><span class="sxs-lookup"><span data-stu-id="7247a-128">A unique ID for the function.</span></span> <span data-ttu-id="7247a-129">Cet ID peut contenir uniquement des points et caractères alphanumériques et ne doit pas être modifié une fois défini.</span><span class="sxs-lookup"><span data-stu-id="7247a-129">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="7247a-130">string</span><span class="sxs-lookup"><span data-stu-id="7247a-130">string</span></span>  |  <span data-ttu-id="7247a-131">Oui</span><span class="sxs-lookup"><span data-stu-id="7247a-131">Yes</span></span>  |  <span data-ttu-id="7247a-132">Nom de la fonction que voient les utilisateurs finaux dans Excel.</span><span class="sxs-lookup"><span data-stu-id="7247a-132">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="7247a-133">Dans Excel, le nom de la fonction sera précédé de l’espace de noms de fonctions personnalisées spécifié dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="7247a-133">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="7247a-134">object</span><span class="sxs-lookup"><span data-stu-id="7247a-134">object</span></span>  |  <span data-ttu-id="7247a-135">Non</span><span class="sxs-lookup"><span data-stu-id="7247a-135">No</span></span>  |  <span data-ttu-id="7247a-136">Vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction.</span><span class="sxs-lookup"><span data-stu-id="7247a-136">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="7247a-137">Reportez-vous aux [options](#options) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="7247a-137">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="7247a-138">tableau</span><span class="sxs-lookup"><span data-stu-id="7247a-138">array</span></span>  |  <span data-ttu-id="7247a-139">Oui</span><span class="sxs-lookup"><span data-stu-id="7247a-139">Yes</span></span>  |  <span data-ttu-id="7247a-140">Tableau qui définit les paramètres d’entrée de la fonction.</span><span class="sxs-lookup"><span data-stu-id="7247a-140">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="7247a-141">Reportez-vous aux [paramètres](#parameters) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="7247a-141">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="7247a-142">objet</span><span class="sxs-lookup"><span data-stu-id="7247a-142">object</span></span>  |  <span data-ttu-id="7247a-143">Oui</span><span class="sxs-lookup"><span data-stu-id="7247a-143">Yes</span></span>  |  <span data-ttu-id="7247a-144">Objet qui définit le type d’informations renvoyées par la fonction.</span><span class="sxs-lookup"><span data-stu-id="7247a-144">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="7247a-145">Reportez-vous au [résultat](#result) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="7247a-145">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="7247a-146">options</span><span class="sxs-lookup"><span data-stu-id="7247a-146">options</span></span>

<span data-ttu-id="7247a-147">L’objet `options` vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction.</span><span class="sxs-lookup"><span data-stu-id="7247a-147">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="7247a-148">Le tableau suivant répertorie les propriétés de l’objet `options`.</span><span class="sxs-lookup"><span data-stu-id="7247a-148">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="7247a-149">Propriété</span><span class="sxs-lookup"><span data-stu-id="7247a-149">Property</span></span>  |  <span data-ttu-id="7247a-150">Type de données</span><span class="sxs-lookup"><span data-stu-id="7247a-150">Data type</span></span>  |  <span data-ttu-id="7247a-151">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="7247a-151">Required</span></span>  |  <span data-ttu-id="7247a-152">Description</span><span class="sxs-lookup"><span data-stu-id="7247a-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="7247a-153">boolean</span><span class="sxs-lookup"><span data-stu-id="7247a-153">boolean</span></span>  |  <span data-ttu-id="7247a-154">Non</span><span class="sxs-lookup"><span data-stu-id="7247a-154">No</span></span><br/><br/><span data-ttu-id="7247a-155">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="7247a-155">Default value is `false`.</span></span>  |  <span data-ttu-id="7247a-156">Si la valeur est `true`, Excel appelle le gestionnaire `onCanceled` chaque fois que l’utilisateur effectue une action ayant pour effet d’annuler la fonction, par exemple, en déclenchant manuellement un recalcul ou en modifiant une cellule référencée par la fonction.</span><span class="sxs-lookup"><span data-stu-id="7247a-156">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="7247a-157">Si vous utilisez cette option, Excel appelle la fonction JavaScript avec un paramètre `caller` supplémentaire</span><span class="sxs-lookup"><span data-stu-id="7247a-157">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="7247a-158">(n’enregistrez ***pas*** ce paramètre dans la propriété `parameters`).</span><span class="sxs-lookup"><span data-stu-id="7247a-158">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="7247a-159">Dans le corps de la fonction, un gestionnaire doit être attribué au membre `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="7247a-159">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="7247a-160">Pour plus d’informations, voir [Annuler une fonction](custom-functions-overview.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="7247a-160">For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="7247a-161">boolean</span><span class="sxs-lookup"><span data-stu-id="7247a-161">boolean</span></span>  |  <span data-ttu-id="7247a-162">Non</span><span class="sxs-lookup"><span data-stu-id="7247a-162">No</span></span><br/><br/><span data-ttu-id="7247a-163">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="7247a-163">Default value is `false`.</span></span>  |  <span data-ttu-id="7247a-164">Si la valeur est `true`, la fonction peut envoyer une sortie à la cellule à plusieurs reprises, même en cas d’appel unique.</span><span class="sxs-lookup"><span data-stu-id="7247a-164">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="7247a-165">Cette option est utile pour des sources de données qui changent rapidement, telles que des valeurs boursières.</span><span class="sxs-lookup"><span data-stu-id="7247a-165">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="7247a-166">Si vous utilisez cette option, Excel appelle la fonction JavaScript avec un paramètre `caller` supplémentaire</span><span class="sxs-lookup"><span data-stu-id="7247a-166">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="7247a-167">(n’enregistrez ***pas*** ce paramètre dans la propriété `parameters`).</span><span class="sxs-lookup"><span data-stu-id="7247a-167">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="7247a-168">La fonction ne doit pas utiliser d’instruction `return`.</span><span class="sxs-lookup"><span data-stu-id="7247a-168">The function should have no `return` statement.</span></span> <span data-ttu-id="7247a-169">Au lieu de cela, la valeur obtenue est transmise en tant qu’argument de la méthode de rappel `caller.setResult`.</span><span class="sxs-lookup"><span data-stu-id="7247a-169">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="7247a-170">Pour plus d’informations, voir [Diffusion en continu de fonctions](custom-functions-overview.md#streaming-functions).</span><span class="sxs-lookup"><span data-stu-id="7247a-170">For more information, see [Streaming functions](custom-functions-overview.md#streaming-functions).</span></span> |
|  `volatile`  | <span data-ttu-id="7247a-171">boolean</span><span class="sxs-lookup"><span data-stu-id="7247a-171">boolean</span></span> | <span data-ttu-id="7247a-172">Non</span><span class="sxs-lookup"><span data-stu-id="7247a-172">No</span></span> <br/><br/><span data-ttu-id="7247a-173">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="7247a-173">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="7247a-174">Si la valeur est `true`, la fonction est recalculée à chaque recalcul d’Excel, et plus à chaque fois que les valeurs dépendantes de la formules sont modifiées.</span><span class="sxs-lookup"><span data-stu-id="7247a-174">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="7247a-175">Une fonction ne peut pas être à la fois diffusée en continu et volatile.</span><span class="sxs-lookup"><span data-stu-id="7247a-175">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="7247a-176">Si les propriétés `stream` et `volatile` sont toutes les deux définies sur `true`, l’option volatile est ignorée.</span><span class="sxs-lookup"><span data-stu-id="7247a-176">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="7247a-177">paramètres</span><span class="sxs-lookup"><span data-stu-id="7247a-177">parameters</span></span>

<span data-ttu-id="7247a-178">La propriété `parameters` est un tableau d’objets paramètre.</span><span class="sxs-lookup"><span data-stu-id="7247a-178">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="7247a-179">Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="7247a-179">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="7247a-180">Propriété</span><span class="sxs-lookup"><span data-stu-id="7247a-180">Property</span></span>  |  <span data-ttu-id="7247a-181">Type de données</span><span class="sxs-lookup"><span data-stu-id="7247a-181">Data type</span></span>  |  <span data-ttu-id="7247a-182">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="7247a-182">Required</span></span>  |  <span data-ttu-id="7247a-183">Description</span><span class="sxs-lookup"><span data-stu-id="7247a-183">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="7247a-184">string</span><span class="sxs-lookup"><span data-stu-id="7247a-184">string</span></span>  |  <span data-ttu-id="7247a-185">Non</span><span class="sxs-lookup"><span data-stu-id="7247a-185">No</span></span> |  <span data-ttu-id="7247a-186">Description du paramètre.</span><span class="sxs-lookup"><span data-stu-id="7247a-186">A description of the parameter.</span></span> <span data-ttu-id="7247a-187">S’affiche dans intelliSense d’Excel.</span><span class="sxs-lookup"><span data-stu-id="7247a-187">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="7247a-188">string</span><span class="sxs-lookup"><span data-stu-id="7247a-188">string</span></span>  |  <span data-ttu-id="7247a-189">Non</span><span class="sxs-lookup"><span data-stu-id="7247a-189">No</span></span>  |  <span data-ttu-id="7247a-190">Doit être **scalaire** (valeur autre que de tableau) ou **matrice** (tableau bidimensionnel).</span><span class="sxs-lookup"><span data-stu-id="7247a-190">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="7247a-191">string</span><span class="sxs-lookup"><span data-stu-id="7247a-191">string</span></span>  |  <span data-ttu-id="7247a-192">Oui</span><span class="sxs-lookup"><span data-stu-id="7247a-192">Yes</span></span>  |  <span data-ttu-id="7247a-193">Le nom du paramètre.</span><span class="sxs-lookup"><span data-stu-id="7247a-193">The name of the parameter.</span></span> <span data-ttu-id="7247a-194">Ce nom s’affiche dans intelliSense d’Excel.</span><span class="sxs-lookup"><span data-stu-id="7247a-194">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="7247a-195">string</span><span class="sxs-lookup"><span data-stu-id="7247a-195">string</span></span>  |  <span data-ttu-id="7247a-196">Non</span><span class="sxs-lookup"><span data-stu-id="7247a-196">No</span></span>  |  <span data-ttu-id="7247a-197">Type de données du paramètre.</span><span class="sxs-lookup"><span data-stu-id="7247a-197">The data type of the parameter.</span></span> <span data-ttu-id="7247a-198">Peut être **boolean**, **number**, **string** ou **any** qui vous permet d’utiliser n’importe lequel des trois types précédents.</span><span class="sxs-lookup"><span data-stu-id="7247a-198">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="7247a-199">Si cette propriété n’est pas spécifiée, le type de données par défaut est **any**.</span><span class="sxs-lookup"><span data-stu-id="7247a-199">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="7247a-200">boolean</span><span class="sxs-lookup"><span data-stu-id="7247a-200">boolean</span></span> | <span data-ttu-id="7247a-201">Non</span><span class="sxs-lookup"><span data-stu-id="7247a-201">No</span></span> | <span data-ttu-id="7247a-202">Si la valeur est `true`, le paramètre est facultatif.</span><span class="sxs-lookup"><span data-stu-id="7247a-202">If `true`, the parameter is optional.</span></span> |

>[!NOTE]
> <span data-ttu-id="7247a-203">Si la propriété `type` d’un paramètre facultatif n’est pas spécifiée ou est définie sur `any`, vous remarquerez peut-être des problèmes tels que des erreurs de linting dans votre IDE et des paramètres facultatifs non affichés lorsque la fonction est saisie dans une cellule Excel.</span><span class="sxs-lookup"><span data-stu-id="7247a-203">If the `type` property of an optional parameter is either not specified or set to `any`, you may notice issues such as linting errors in your IDE and optional parameters not being displayed when the function is being entered into a cell in Excel.</span></span> <span data-ttu-id="7247a-204">Ces problèmes seront résolus en décembre 2018.</span><span class="sxs-lookup"><span data-stu-id="7247a-204">This is projected to change in December of 2018.</span></span>

## <a name="result"></a><span data-ttu-id="7247a-205">résultat</span><span class="sxs-lookup"><span data-stu-id="7247a-205">result</span></span>

<span data-ttu-id="7247a-206">L’objet `result` définit le type des informations renvoyées par la fonction.</span><span class="sxs-lookup"><span data-stu-id="7247a-206">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="7247a-207">Le tableau suivant répertorie les propriétés de l’objet `result`.</span><span class="sxs-lookup"><span data-stu-id="7247a-207">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="7247a-208">Propriété</span><span class="sxs-lookup"><span data-stu-id="7247a-208">Property</span></span>  |  <span data-ttu-id="7247a-209">Type de données</span><span class="sxs-lookup"><span data-stu-id="7247a-209">Data type</span></span>  |  <span data-ttu-id="7247a-210">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="7247a-210">Required</span></span>  |  <span data-ttu-id="7247a-211">Description</span><span class="sxs-lookup"><span data-stu-id="7247a-211">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="7247a-212">string</span><span class="sxs-lookup"><span data-stu-id="7247a-212">string</span></span>  |  <span data-ttu-id="7247a-213">Non</span><span class="sxs-lookup"><span data-stu-id="7247a-213">No</span></span>  |  <span data-ttu-id="7247a-214">Doit être **scalaire** (valeur autre que de tableau) ou **matrice** (tableau bidimensionnel).</span><span class="sxs-lookup"><span data-stu-id="7247a-214">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="7247a-215">string</span><span class="sxs-lookup"><span data-stu-id="7247a-215">string</span></span>  |  <span data-ttu-id="7247a-216">Oui</span><span class="sxs-lookup"><span data-stu-id="7247a-216">Yes</span></span>  |  <span data-ttu-id="7247a-217">Type de données du paramètre.</span><span class="sxs-lookup"><span data-stu-id="7247a-217">The data type of the parameter.</span></span> <span data-ttu-id="7247a-218">Doit être **boolean**, **number**, **string** ou **any** qui vous permet d’utiliser n’importe lequel des trois types précédents.</span><span class="sxs-lookup"><span data-stu-id="7247a-218">Must be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> |

## <a name="see-also"></a><span data-ttu-id="7247a-219">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="7247a-219">See also</span></span>

* [<span data-ttu-id="7247a-220">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="7247a-220">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="7247a-221">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="7247a-221">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="7247a-222">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="7247a-222">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="7247a-223">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="7247a-223">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
