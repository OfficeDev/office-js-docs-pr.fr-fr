---
ms.date: 06/20/2019
description: Définissez des métadonnées pour des fonctions personnalisées dans Excel.
title: Métadonnées pour les fonctions personnalisées dans Excel
localization_priority: Normal
ms.openlocfilehash: a9fbefb7ea1c5474d26b668d3a4f64ed68ae36f7
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454635"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="cfea1-103">Métadonnées des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="cfea1-103">Custom functions metadata</span></span>

<span data-ttu-id="cfea1-104">Lorsque vous définissez des [fonctions personnalisées](custom-functions-overview.md) dans votre complément Excel, votre projet de complément inclut un fichier de métadonnées JSON qui fournit les informations dont Excel a besoin pour enregistrer les fonctions personnalisées et les rendre accessibles aux utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="cfea1-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project includes a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="cfea1-105">Ce fichier est généré:</span><span class="sxs-lookup"><span data-stu-id="cfea1-105">This file is generated either:</span></span>

- <span data-ttu-id="cfea1-106">Par vous-même, dans un fichier JSON manuscrit</span><span class="sxs-lookup"><span data-stu-id="cfea1-106">By you, in a handwritten JSON file</span></span>
- <span data-ttu-id="cfea1-107">À partir des commentaires JSDoc que vous entrez au début de votre fonction</span><span class="sxs-lookup"><span data-stu-id="cfea1-107">From the JSDoc comments you enter at the beginning of your function</span></span>

<span data-ttu-id="cfea1-108">Les fonctions personnalisées sont inscrites lorsque l’utilisateur exécute le complément pour la première fois et après qu’il est disponible pour le même utilisateur dans tous les classeurs.</span><span class="sxs-lookup"><span data-stu-id="cfea1-108">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

<span data-ttu-id="cfea1-109">Cet article décrit le format du fichier de métadonnées JSON, en supposant que vous l’écrivez manuellement.</span><span class="sxs-lookup"><span data-stu-id="cfea1-109">This article describes the format of the JSON metadata file, assuming you are writing it by hand.</span></span> <span data-ttu-id="cfea1-110">Pour plus d’informations sur la génération de fichiers JSON de commentaire JSDoc, voir [generate JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="cfea1-110">For information about JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="cfea1-111">Pour plus d’informations sur les autres fichiers à inclure dans votre projet de complément afin d’activer les fonctions personnalisées, voir [Créer des fonctions personnalisées dans Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="cfea1-111">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

<span data-ttu-id="cfea1-112">Les paramètres serveur sur le serveur qui héberge le fichier JSON doivent avoir [cors](https://developer.mozilla.org/docs/Web/HTTP/CORS) activé afin que les fonctions personnalisées fonctionnent correctement dans Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="cfea1-112">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel on the web.</span></span>

## <a name="example-metadata"></a><span data-ttu-id="cfea1-113">Exemple de métadonnées</span><span class="sxs-lookup"><span data-stu-id="cfea1-113">Example metadata</span></span>

<span data-ttu-id="cfea1-114">L’exemple suivant montre le contenu d’un fichier de métadonnées JSON pour un complément qui définit des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="cfea1-114">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="cfea1-115">Les sections qui suivent cet exemple fournissent des informations détaillées sur les propriétés individuelles au sein de cet exemple JSON.</span><span class="sxs-lookup"><span data-stu-id="cfea1-115">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
        "dimensionality": "scalar"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE", 
      "description":  "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
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
> <span data-ttu-id="cfea1-116">Un exemple de fichier JSON complet est disponible dans l’historique de validation du référentiel [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) github.</span><span class="sxs-lookup"><span data-stu-id="cfea1-116">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="cfea1-117">Lorsque le projet a été ajusté pour générer automatiquement JSON, un échantillon complet de JSON manuscrit est uniquement disponible dans les versions précédentes du projet.</span><span class="sxs-lookup"><span data-stu-id="cfea1-117">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="functions"></a><span data-ttu-id="cfea1-118">fonctions</span><span class="sxs-lookup"><span data-stu-id="cfea1-118">functions</span></span> 

<span data-ttu-id="cfea1-119">La propriété `functions` est un tableau d’objets de fonction personnalisés.</span><span class="sxs-lookup"><span data-stu-id="cfea1-119">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="cfea1-120">Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="cfea1-120">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="cfea1-121">Propriété</span><span class="sxs-lookup"><span data-stu-id="cfea1-121">Property</span></span>  |  <span data-ttu-id="cfea1-122">Type de données</span><span class="sxs-lookup"><span data-stu-id="cfea1-122">Data type</span></span>  |  <span data-ttu-id="cfea1-123">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="cfea1-123">Required</span></span>  |  <span data-ttu-id="cfea1-124">Description</span><span class="sxs-lookup"><span data-stu-id="cfea1-124">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="cfea1-125">string</span><span class="sxs-lookup"><span data-stu-id="cfea1-125">string</span></span>  |  <span data-ttu-id="cfea1-126">Non</span><span class="sxs-lookup"><span data-stu-id="cfea1-126">No</span></span>  |  <span data-ttu-id="cfea1-127">Description de la fonction que voient les utilisateurs finaux dans Excel.</span><span class="sxs-lookup"><span data-stu-id="cfea1-127">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="cfea1-128">Par exemple, **convertit une valeur Celsius en valeur Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="cfea1-128">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="cfea1-129">string</span><span class="sxs-lookup"><span data-stu-id="cfea1-129">string</span></span>  |   <span data-ttu-id="cfea1-130">Non</span><span class="sxs-lookup"><span data-stu-id="cfea1-130">No</span></span>  |  <span data-ttu-id="cfea1-131">URL fournissant des informations sur la fonction</span><span class="sxs-lookup"><span data-stu-id="cfea1-131">URL that provides information about the function.</span></span> <span data-ttu-id="cfea1-132">(elle est affichée dans un volet des tâches). Par exemple, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span><span class="sxs-lookup"><span data-stu-id="cfea1-132">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span> |
| `id`     | <span data-ttu-id="cfea1-133">string</span><span class="sxs-lookup"><span data-stu-id="cfea1-133">string</span></span> | <span data-ttu-id="cfea1-134">Oui</span><span class="sxs-lookup"><span data-stu-id="cfea1-134">Yes</span></span> | <span data-ttu-id="cfea1-135">Un ID unique pour la fonction.</span><span class="sxs-lookup"><span data-stu-id="cfea1-135">A unique ID for the function.</span></span> <span data-ttu-id="cfea1-136">Cet ID peut contenir uniquement des points et caractères alphanumériques et ne doit pas être modifié une fois défini.</span><span class="sxs-lookup"><span data-stu-id="cfea1-136">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="cfea1-137">string</span><span class="sxs-lookup"><span data-stu-id="cfea1-137">string</span></span>  |  <span data-ttu-id="cfea1-138">Oui</span><span class="sxs-lookup"><span data-stu-id="cfea1-138">Yes</span></span>  |  <span data-ttu-id="cfea1-139">Nom de la fonction que voient les utilisateurs finaux dans Excel.</span><span class="sxs-lookup"><span data-stu-id="cfea1-139">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="cfea1-140">Dans Excel, le nom de la fonction sera précédé de l’espace de noms de fonctions personnalisées spécifié dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="cfea1-140">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="cfea1-141">object</span><span class="sxs-lookup"><span data-stu-id="cfea1-141">object</span></span>  |  <span data-ttu-id="cfea1-142">Non</span><span class="sxs-lookup"><span data-stu-id="cfea1-142">No</span></span>  |  <span data-ttu-id="cfea1-143">Vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction.</span><span class="sxs-lookup"><span data-stu-id="cfea1-143">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="cfea1-144">Reportez-vous aux [options](#options) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="cfea1-144">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="cfea1-145">tableau</span><span class="sxs-lookup"><span data-stu-id="cfea1-145">array</span></span>  |  <span data-ttu-id="cfea1-146">Oui</span><span class="sxs-lookup"><span data-stu-id="cfea1-146">Yes</span></span>  |  <span data-ttu-id="cfea1-147">Tableau qui définit les paramètres d’entrée de la fonction.</span><span class="sxs-lookup"><span data-stu-id="cfea1-147">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="cfea1-148">Reportez-vous aux [paramètres](#parameters) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="cfea1-148">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="cfea1-149">objet</span><span class="sxs-lookup"><span data-stu-id="cfea1-149">object</span></span>  |  <span data-ttu-id="cfea1-150">Oui</span><span class="sxs-lookup"><span data-stu-id="cfea1-150">Yes</span></span>  |  <span data-ttu-id="cfea1-151">Objet qui définit le type d’informations renvoyées par la fonction.</span><span class="sxs-lookup"><span data-stu-id="cfea1-151">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="cfea1-152">Reportez-vous au [résultat](#result) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="cfea1-152">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="cfea1-153">options</span><span class="sxs-lookup"><span data-stu-id="cfea1-153">options</span></span>

<span data-ttu-id="cfea1-154">L’objet `options` vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction.</span><span class="sxs-lookup"><span data-stu-id="cfea1-154">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="cfea1-155">Le tableau suivant répertorie les propriétés de l’objet `options`.</span><span class="sxs-lookup"><span data-stu-id="cfea1-155">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="cfea1-156">Propriété</span><span class="sxs-lookup"><span data-stu-id="cfea1-156">Property</span></span>  |  <span data-ttu-id="cfea1-157">Type de données</span><span class="sxs-lookup"><span data-stu-id="cfea1-157">Data type</span></span>  |  <span data-ttu-id="cfea1-158">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="cfea1-158">Required</span></span>  |  <span data-ttu-id="cfea1-159">Description</span><span class="sxs-lookup"><span data-stu-id="cfea1-159">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="cfea1-160">boolean</span><span class="sxs-lookup"><span data-stu-id="cfea1-160">boolean</span></span>  |  <span data-ttu-id="cfea1-161">Non</span><span class="sxs-lookup"><span data-stu-id="cfea1-161">No</span></span><br/><br/><span data-ttu-id="cfea1-162">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="cfea1-162">Default value is `false`.</span></span>  |  <span data-ttu-id="cfea1-163">Si la valeur est `true`, Excel appelle le gestionnaire `CancelableInvocation` chaque fois que l’utilisateur effectue une action ayant pour effet d’annuler la fonction, par exemple, en déclenchant manuellement un recalcul ou en modifiant une cellule référencée par la fonction.</span><span class="sxs-lookup"><span data-stu-id="cfea1-163">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="cfea1-164">Les fonctions annulables sont généralement utilisées uniquement pour les fonctions asynchrones qui renvoient un seul résultat et doivent gérer l’annulation d’une demande de données.</span><span class="sxs-lookup"><span data-stu-id="cfea1-164">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="cfea1-165">Une fonction ne peut pas être à la fois en continu et annulable.</span><span class="sxs-lookup"><span data-stu-id="cfea1-165">A function cannot be both streaming and cancelable.</span></span> <span data-ttu-id="cfea1-166">Pour plus d’informations, reportez-vous à la remarque à la fin de la [création d’une fonction de diffusion en continu](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="cfea1-166">For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
|  `requiresAddress`  | <span data-ttu-id="cfea1-167">boolean</span><span class="sxs-lookup"><span data-stu-id="cfea1-167">boolean</span></span> | <span data-ttu-id="cfea1-168">Non</span><span class="sxs-lookup"><span data-stu-id="cfea1-168">No</span></span> <br/><br/><span data-ttu-id="cfea1-169">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="cfea1-169">Default value is `false`.</span></span> | <span data-ttu-id="cfea1-170">Si `true`votre fonction personnalisée peut accéder à l’adresse de la cellule qui a appelé votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="cfea1-170">If `true`, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="cfea1-171">Pour obtenir l’adresse de la cellule qui a appelé votre fonction personnalisée, utilisez Context. Address dans votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="cfea1-171">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="cfea1-172">Pour plus d’informations, consultez la rubrique relative au [paramètre context de la cellule](/office/dev/add-ins/excel/custom-functions-parameter-options#addressing-cells-context-parameter)Addressing.</span><span class="sxs-lookup"><span data-stu-id="cfea1-172">For more information, see [Addressing cell's context parameter](/office/dev/add-ins/excel/custom-functions-parameter-options#addressing-cells-context-parameter).</span></span> <span data-ttu-id="cfea1-173">Les fonctions personnalisées ne peuvent pas être définies à la fois en diffusion en continu et requiresAddress.</span><span class="sxs-lookup"><span data-stu-id="cfea1-173">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="cfea1-174">Lorsque vous utilisez cette option, le paramètre «invocation» doit être le dernier paramètre passé dans options.</span><span class="sxs-lookup"><span data-stu-id="cfea1-174">When using this option, the 'invocation' parameter must be the last parameter passed in options.</span></span> |
|  `stream`  |  <span data-ttu-id="cfea1-175">boolean</span><span class="sxs-lookup"><span data-stu-id="cfea1-175">boolean</span></span>  |  <span data-ttu-id="cfea1-176">Non</span><span class="sxs-lookup"><span data-stu-id="cfea1-176">No</span></span><br/><br/><span data-ttu-id="cfea1-177">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="cfea1-177">Default value is `false`.</span></span>  |  <span data-ttu-id="cfea1-178">Si la valeur est `true`, la fonction peut envoyer une sortie à la cellule à plusieurs reprises, même en cas d’appel unique.</span><span class="sxs-lookup"><span data-stu-id="cfea1-178">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="cfea1-179">Cette option est utile pour des sources de données qui changent rapidement, telles que des valeurs boursières.</span><span class="sxs-lookup"><span data-stu-id="cfea1-179">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="cfea1-180">La fonction ne doit pas utiliser d’instruction `return`.</span><span class="sxs-lookup"><span data-stu-id="cfea1-180">The function should have no `return` statement.</span></span> <span data-ttu-id="cfea1-181">Au lieu de cela, la valeur obtenue est transmise en tant qu’argument de la méthode de rappel `StreamingInvocation.setResult`.</span><span class="sxs-lookup"><span data-stu-id="cfea1-181">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="cfea1-182">Pour plus d’informations, voir [Diffusion en continu de fonctions](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="cfea1-182">For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
|  `volatile`  | <span data-ttu-id="cfea1-183">boolean</span><span class="sxs-lookup"><span data-stu-id="cfea1-183">boolean</span></span> | <span data-ttu-id="cfea1-184">Non</span><span class="sxs-lookup"><span data-stu-id="cfea1-184">No</span></span> <br/><br/><span data-ttu-id="cfea1-185">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="cfea1-185">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="cfea1-186">Si la valeur est `true`, la fonction est recalculée à chaque recalcul d’Excel, et plus à chaque fois que les valeurs dépendantes de la formules sont modifiées.</span><span class="sxs-lookup"><span data-stu-id="cfea1-186">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="cfea1-187">Une fonction ne peut pas être à la fois diffusée en continu et volatile.</span><span class="sxs-lookup"><span data-stu-id="cfea1-187">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="cfea1-188">Si les propriétés `stream` et `volatile` sont toutes les deux définies sur `true`, l’option volatile est ignorée.</span><span class="sxs-lookup"><span data-stu-id="cfea1-188">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="cfea1-189">paramètres</span><span class="sxs-lookup"><span data-stu-id="cfea1-189">parameters</span></span>

<span data-ttu-id="cfea1-190">La propriété `parameters` est un tableau d’objets paramètre.</span><span class="sxs-lookup"><span data-stu-id="cfea1-190">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="cfea1-191">Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="cfea1-191">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="cfea1-192">Propriété</span><span class="sxs-lookup"><span data-stu-id="cfea1-192">Property</span></span>  |  <span data-ttu-id="cfea1-193">Type de données</span><span class="sxs-lookup"><span data-stu-id="cfea1-193">Data type</span></span>  |  <span data-ttu-id="cfea1-194">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="cfea1-194">Required</span></span>  |  <span data-ttu-id="cfea1-195">Description</span><span class="sxs-lookup"><span data-stu-id="cfea1-195">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="cfea1-196">string</span><span class="sxs-lookup"><span data-stu-id="cfea1-196">string</span></span>  |  <span data-ttu-id="cfea1-197">Non</span><span class="sxs-lookup"><span data-stu-id="cfea1-197">No</span></span> |  <span data-ttu-id="cfea1-198">Description du paramètre.</span><span class="sxs-lookup"><span data-stu-id="cfea1-198">A description of the parameter.</span></span> <span data-ttu-id="cfea1-199">S’affiche dans intelliSense d’Excel.</span><span class="sxs-lookup"><span data-stu-id="cfea1-199">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="cfea1-200">string</span><span class="sxs-lookup"><span data-stu-id="cfea1-200">string</span></span>  |  <span data-ttu-id="cfea1-201">Non</span><span class="sxs-lookup"><span data-stu-id="cfea1-201">No</span></span>  |  <span data-ttu-id="cfea1-202">Doit être **scalaire** (valeur autre que de tableau) ou **matrice** (tableau bidimensionnel).</span><span class="sxs-lookup"><span data-stu-id="cfea1-202">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="cfea1-203">string</span><span class="sxs-lookup"><span data-stu-id="cfea1-203">string</span></span>  |  <span data-ttu-id="cfea1-204">Oui</span><span class="sxs-lookup"><span data-stu-id="cfea1-204">Yes</span></span>  |  <span data-ttu-id="cfea1-205">Le nom du paramètre.</span><span class="sxs-lookup"><span data-stu-id="cfea1-205">The name of the parameter.</span></span> <span data-ttu-id="cfea1-206">Ce nom s’affiche dans intelliSense d’Excel.</span><span class="sxs-lookup"><span data-stu-id="cfea1-206">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="cfea1-207">string</span><span class="sxs-lookup"><span data-stu-id="cfea1-207">string</span></span>  |  <span data-ttu-id="cfea1-208">Non</span><span class="sxs-lookup"><span data-stu-id="cfea1-208">No</span></span>  |  <span data-ttu-id="cfea1-209">Type de données du paramètre.</span><span class="sxs-lookup"><span data-stu-id="cfea1-209">The data type of the parameter.</span></span> <span data-ttu-id="cfea1-210">Peut être **boolean**, **number**, **string** ou **any** qui vous permet d’utiliser n’importe lequel des trois types précédents.</span><span class="sxs-lookup"><span data-stu-id="cfea1-210">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="cfea1-211">Si cette propriété n’est pas spécifiée, le type de données par défaut est **any**.</span><span class="sxs-lookup"><span data-stu-id="cfea1-211">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="cfea1-212">boolean</span><span class="sxs-lookup"><span data-stu-id="cfea1-212">boolean</span></span> | <span data-ttu-id="cfea1-213">Non</span><span class="sxs-lookup"><span data-stu-id="cfea1-213">No</span></span> | <span data-ttu-id="cfea1-214">Si la valeur est `true`, le paramètre est facultatif.</span><span class="sxs-lookup"><span data-stu-id="cfea1-214">If `true`, the parameter is optional.</span></span> |

## <a name="result"></a><span data-ttu-id="cfea1-215">résultat</span><span class="sxs-lookup"><span data-stu-id="cfea1-215">result</span></span>

<span data-ttu-id="cfea1-216">L’objet `result` définit le type des informations renvoyées par la fonction.</span><span class="sxs-lookup"><span data-stu-id="cfea1-216">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="cfea1-217">Le tableau suivant répertorie les propriétés de l’objet `result`.</span><span class="sxs-lookup"><span data-stu-id="cfea1-217">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="cfea1-218">Propriété</span><span class="sxs-lookup"><span data-stu-id="cfea1-218">Property</span></span>  |  <span data-ttu-id="cfea1-219">Type de données</span><span class="sxs-lookup"><span data-stu-id="cfea1-219">Data type</span></span>  |  <span data-ttu-id="cfea1-220">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="cfea1-220">Required</span></span>  |  <span data-ttu-id="cfea1-221">Description</span><span class="sxs-lookup"><span data-stu-id="cfea1-221">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="cfea1-222">string</span><span class="sxs-lookup"><span data-stu-id="cfea1-222">string</span></span>  |  <span data-ttu-id="cfea1-223">Non</span><span class="sxs-lookup"><span data-stu-id="cfea1-223">No</span></span>  |  <span data-ttu-id="cfea1-224">Doit être **scalaire** (valeur autre que de tableau) ou **matrice** (tableau bidimensionnel).</span><span class="sxs-lookup"><span data-stu-id="cfea1-224">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="next-steps"></a><span data-ttu-id="cfea1-225">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="cfea1-225">Next steps</span></span>
<span data-ttu-id="cfea1-226">Découvrez les [meilleures pratiques de dénomination de votre fonction](custom-functions-naming.md) ou Découvrez comment [localiser votre fonction](custom-functions-localize.md) à l’aide de la méthode JSON manuscrite décrite précédemment.</span><span class="sxs-lookup"><span data-stu-id="cfea1-226">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="cfea1-227">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="cfea1-227">See also</span></span>

* [<span data-ttu-id="cfea1-228">Générer automatiquement des métadonnées JSON pour des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="cfea1-228">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="cfea1-229">Options des paramètres de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="cfea1-229">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
* [<span data-ttu-id="cfea1-230">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="cfea1-230">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="cfea1-231">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="cfea1-231">Create custom functions in Excel</span></span>](custom-functions-overview.md)
