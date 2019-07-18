---
ms.date: 07/15/2019
description: Définissez des métadonnées JSON pour les fonctions personnalisées dans Excel et associez vos ID de fonction et propriétés de nom.
title: Métadonnées pour les fonctions personnalisées dans Excel
localization_priority: Normal
ms.openlocfilehash: b0e015cfa439651420487db4885647f5c7de7da8
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771560"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="fae4b-103">Métadonnées des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="fae4b-103">Custom functions metadata</span></span>

<span data-ttu-id="fae4b-104">Comme décrit dans l’article [vue d’ensemble des fonctions personnalisées](custom-functions-overview.md) , un projet de fonctions personnalisées doit inclure un fichier de métadonnées JSON et un fichier script (JavaScript ou machine à écriture) pour enregistrer une fonction, le rendant ainsi disponible.</span><span class="sxs-lookup"><span data-stu-id="fae4b-104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to register a function, making it available for use.</span></span> <span data-ttu-id="fae4b-105">Les fonctions personnalisées sont inscrites lorsque l’utilisateur exécute le complément pour la première fois et après qu’il est disponible pour le même utilisateur dans tous les classeurs.</span><span class="sxs-lookup"><span data-stu-id="fae4b-105">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="fae4b-106">Il est recommandé d’utiliser la génération automatique JSON dans la mesure du possible `yo office` , à l’aide des fichiers de l’échafaudage, de la même manière que le processus illustré dans le [didacticiel de fonction personnalisée Excel](../tutorials/excel-tutorial-create-custom-functions.md) , car ce processus est plus facile et moins sujet aux erreurs de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="fae4b-106">It is recommended that you use JSON autogeneration when possible, using the `yo office` scaffold files, similar to the process shown in the [Excel Custom Function tutorial](../tutorials/excel-tutorial-create-custom-functions.md) because this process is easier and less prone to user error.</span></span> <span data-ttu-id="fae4b-107">Pour plus d’informations sur le processus de génération de fichiers JSON de commentaire JSDoc, voir [génération de métadonnées JSON pour les fonctions personnalisées](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="fae4b-107">For more information on the process of JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="fae4b-108">Toutefois, vous pouvez créer un projet de fonctions personnalisées de a à z; pour ce faire, vous devez:</span><span class="sxs-lookup"><span data-stu-id="fae4b-108">However, you can make a custom functions project from scratch; it requires that you:</span></span>

- <span data-ttu-id="fae4b-109">Écrire votre fichier JSON manuellement</span><span class="sxs-lookup"><span data-stu-id="fae4b-109">Write your JSON file by hand</span></span>
- <span data-ttu-id="fae4b-110">Vérifier que votre fichier manifeste est connecté à votre fichier JSON créé manuellement</span><span class="sxs-lookup"><span data-stu-id="fae4b-110">Check that your manifest file is connected to your hand-authored JSON file</span></span>
- <span data-ttu-id="fae4b-111">Associez les fonctions `id` et `name` les propriétés dans le fichier de script pour enregistrer vos fonctions</span><span class="sxs-lookup"><span data-stu-id="fae4b-111">Associate your functions' `id` and `name` properties in the script file in order to register your functions</span></span>

<span data-ttu-id="fae4b-112">Cet article vous explique comment effectuer ces trois étapes.</span><span class="sxs-lookup"><span data-stu-id="fae4b-112">This article will show you how to do all three of these steps.</span></span>

> [!NOTE]
> <span data-ttu-id="fae4b-113">Contrairement aux fichiers de `yo office` l’échafaudage, vous devez raccorder votre manifeste au fichier JSON que vous créez, via la `<Resources>` section de votre fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="fae4b-113">In contrast with the `yo office` scaffold files, you need to hook your manifest up with the JSON file you create, through the `<Resources>` section in your XML manifest file.</span></span> <span data-ttu-id="fae4b-114">Notez que [cors](https://developer.mozilla.org/docs/Web/HTTP/CORS) doit être activé pour les paramètres serveur sur le serveur qui héberge le fichier JSON afin que les fonctions personnalisées fonctionnent correctement dans Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="fae4b-114">Note that the server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel on the web.</span></span>

## <a name="authoring-metadata-and-hooking-up-to-the-manifest"></a><span data-ttu-id="fae4b-115">Création de métadonnées et raccordement au manifeste</span><span class="sxs-lookup"><span data-stu-id="fae4b-115">Authoring metadata and hooking up to the manifest</span></span>

<span data-ttu-id="fae4b-116">Vous devez créer un fichier JSON dans votre projet et fournir toutes les informations sur les fonctions qu’il contient, telles que les paramètres de la fonction.</span><span class="sxs-lookup"><span data-stu-id="fae4b-116">You need to create a JSON file in your project and provide all the details about your functions in it, such as the function's parameters.</span></span> <span data-ttu-id="fae4b-117">Consultez l' [exemple de métadonnées suivant](#json-metadata-example) et [la référence](#metadata-reference) de métadonnées pour obtenir la liste complète des propriétés de fonction.</span><span class="sxs-lookup"><span data-stu-id="fae4b-117">See the [following metadata example](#json-metadata-example) and [the metadata reference](#metadata-reference) for a complete list of function properties.</span></span>

<span data-ttu-id="fae4b-118">Vous devez également vous assurer que votre fichier manifeste XML fait référence à votre fichier JSON `<Resources>` dans la section, comme dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="fae4b-118">You also need to make sure your XML manifest file references your JSON file in the `<Resources>` section, similar to the following example.</span></span>

```json
<Resources>
    <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
            <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
    </bt:ShortStrings>
</Resources>
```

## <a name="json-metadata-example"></a><span data-ttu-id="fae4b-119">Exemple de métadonnées JSON</span><span class="sxs-lookup"><span data-stu-id="fae4b-119">JSON metadata example</span></span>

<span data-ttu-id="fae4b-120">L’exemple suivant montre le contenu d’un fichier de métadonnées JSON pour un complément qui définit des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="fae4b-120">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="fae4b-121">Les sections qui suivent cet exemple fournissent des informations détaillées sur les propriétés individuelles au sein de cet exemple JSON.</span><span class="sxs-lookup"><span data-stu-id="fae4b-121">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
      "description": "Count up from zero",
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
      "description": "Get the second highest number from a range",
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
> <span data-ttu-id="fae4b-122">Un exemple de fichier JSON complet est disponible dans l’historique de validation du référentiel [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) github.</span><span class="sxs-lookup"><span data-stu-id="fae4b-122">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="fae4b-123">Lorsque le projet a été ajusté pour générer automatiquement JSON, un échantillon complet de JSON manuscrit est uniquement disponible dans les versions précédentes du projet.</span><span class="sxs-lookup"><span data-stu-id="fae4b-123">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="metadata-reference"></a><span data-ttu-id="fae4b-124">Référence de métadonnées</span><span class="sxs-lookup"><span data-stu-id="fae4b-124">Metadata reference</span></span>

### <a name="functions"></a><span data-ttu-id="fae4b-125">fonctions</span><span class="sxs-lookup"><span data-stu-id="fae4b-125">functions</span></span>

<span data-ttu-id="fae4b-126">La propriété `functions` est un tableau d’objets de fonction personnalisés.</span><span class="sxs-lookup"><span data-stu-id="fae4b-126">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="fae4b-127">Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="fae4b-127">The following table lists the properties of each object.</span></span>

| <span data-ttu-id="fae4b-128">Propriété</span><span class="sxs-lookup"><span data-stu-id="fae4b-128">Property</span></span>      | <span data-ttu-id="fae4b-129">Type de données</span><span class="sxs-lookup"><span data-stu-id="fae4b-129">Data type</span></span> | <span data-ttu-id="fae4b-130">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="fae4b-130">Required</span></span> | <span data-ttu-id="fae4b-131">Description</span><span class="sxs-lookup"><span data-stu-id="fae4b-131">Description</span></span>                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | <span data-ttu-id="fae4b-132">string</span><span class="sxs-lookup"><span data-stu-id="fae4b-132">string</span></span>    | <span data-ttu-id="fae4b-133">Non</span><span class="sxs-lookup"><span data-stu-id="fae4b-133">No</span></span>       | <span data-ttu-id="fae4b-134">Description de la fonction que voient les utilisateurs finaux dans Excel.</span><span class="sxs-lookup"><span data-stu-id="fae4b-134">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="fae4b-135">Par exemple, **convertit une valeur Celsius en valeur Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="fae4b-135">For example, **Converts a Celsius value to Fahrenheit**.</span></span>                                                            |
| `helpUrl`     | <span data-ttu-id="fae4b-136">string</span><span class="sxs-lookup"><span data-stu-id="fae4b-136">string</span></span>    | <span data-ttu-id="fae4b-137">Non</span><span class="sxs-lookup"><span data-stu-id="fae4b-137">No</span></span>       | <span data-ttu-id="fae4b-138">URL fournissant des informations sur la fonction</span><span class="sxs-lookup"><span data-stu-id="fae4b-138">URL that provides information about the function.</span></span> <span data-ttu-id="fae4b-139">(elle est affichée dans un volet des tâches). Par exemple, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span><span class="sxs-lookup"><span data-stu-id="fae4b-139">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span>                      |
| `id`          | <span data-ttu-id="fae4b-140">string</span><span class="sxs-lookup"><span data-stu-id="fae4b-140">string</span></span>    | <span data-ttu-id="fae4b-141">Oui</span><span class="sxs-lookup"><span data-stu-id="fae4b-141">Yes</span></span>      | <span data-ttu-id="fae4b-142">Un ID unique pour la fonction.</span><span class="sxs-lookup"><span data-stu-id="fae4b-142">A unique ID for the function.</span></span> <span data-ttu-id="fae4b-143">Cet ID peut contenir uniquement des points et caractères alphanumériques et ne doit pas être modifié une fois défini.</span><span class="sxs-lookup"><span data-stu-id="fae4b-143">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span>                                            |
| `name`        | <span data-ttu-id="fae4b-144">string</span><span class="sxs-lookup"><span data-stu-id="fae4b-144">string</span></span>    | <span data-ttu-id="fae4b-145">Oui</span><span class="sxs-lookup"><span data-stu-id="fae4b-145">Yes</span></span>      | <span data-ttu-id="fae4b-146">Nom de la fonction que voient les utilisateurs finaux dans Excel.</span><span class="sxs-lookup"><span data-stu-id="fae4b-146">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="fae4b-147">Dans Excel, le nom de la fonction sera précédé de l’espace de noms de fonctions personnalisées spécifié dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="fae4b-147">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `options`     | <span data-ttu-id="fae4b-148">object</span><span class="sxs-lookup"><span data-stu-id="fae4b-148">object</span></span>    | <span data-ttu-id="fae4b-149">Non</span><span class="sxs-lookup"><span data-stu-id="fae4b-149">No</span></span>       | <span data-ttu-id="fae4b-150">Vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction.</span><span class="sxs-lookup"><span data-stu-id="fae4b-150">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="fae4b-151">Reportez-vous aux [options](#options) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="fae4b-151">See [options](#options) for details.</span></span>                                                          |
| `parameters`  | <span data-ttu-id="fae4b-152">tableau</span><span class="sxs-lookup"><span data-stu-id="fae4b-152">array</span></span>     | <span data-ttu-id="fae4b-153">Oui</span><span class="sxs-lookup"><span data-stu-id="fae4b-153">Yes</span></span>      | <span data-ttu-id="fae4b-154">Tableau qui définit les paramètres d’entrée de la fonction.</span><span class="sxs-lookup"><span data-stu-id="fae4b-154">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="fae4b-155">Pour plus d’informations, consultez la rubrique [paramètres](#parameters) .</span><span class="sxs-lookup"><span data-stu-id="fae4b-155">See [parameters](#parameters) for details.</span></span>                                                                             |
| `result`      | <span data-ttu-id="fae4b-156">objet</span><span class="sxs-lookup"><span data-stu-id="fae4b-156">object</span></span>    | <span data-ttu-id="fae4b-157">Oui</span><span class="sxs-lookup"><span data-stu-id="fae4b-157">Yes</span></span>      | <span data-ttu-id="fae4b-158">Objet qui définit le type d’informations renvoyées par la fonction.</span><span class="sxs-lookup"><span data-stu-id="fae4b-158">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="fae4b-159">Reportez-vous au [résultat](#result) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="fae4b-159">See [result](#result) for details.</span></span>                                                                 |

### <a name="options"></a><span data-ttu-id="fae4b-160">options</span><span class="sxs-lookup"><span data-stu-id="fae4b-160">options</span></span>

<span data-ttu-id="fae4b-161">L’objet `options` vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction.</span><span class="sxs-lookup"><span data-stu-id="fae4b-161">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="fae4b-162">Le tableau suivant répertorie les propriétés de l’objet `options`.</span><span class="sxs-lookup"><span data-stu-id="fae4b-162">The following table lists the properties of the `options` object.</span></span>

| <span data-ttu-id="fae4b-163">Propriété</span><span class="sxs-lookup"><span data-stu-id="fae4b-163">Property</span></span>          | <span data-ttu-id="fae4b-164">Type de données</span><span class="sxs-lookup"><span data-stu-id="fae4b-164">Data type</span></span> | <span data-ttu-id="fae4b-165">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="fae4b-165">Required</span></span>                               | <span data-ttu-id="fae4b-166">Description</span><span class="sxs-lookup"><span data-stu-id="fae4b-166">Description</span></span>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| :---------------- | :-------- | :------------------------------------- | :--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `cancelable`      | <span data-ttu-id="fae4b-167">boolean</span><span class="sxs-lookup"><span data-stu-id="fae4b-167">boolean</span></span>   | <span data-ttu-id="fae4b-168">Non</span><span class="sxs-lookup"><span data-stu-id="fae4b-168">No</span></span><br/><br/><span data-ttu-id="fae4b-169">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="fae4b-169">Default value is `false`.</span></span>  | <span data-ttu-id="fae4b-170">Si la valeur est `true`, Excel appelle le gestionnaire `CancelableInvocation` chaque fois que l’utilisateur effectue une action ayant pour effet d’annuler la fonction, par exemple, en déclenchant manuellement un recalcul ou en modifiant une cellule référencée par la fonction.</span><span class="sxs-lookup"><span data-stu-id="fae4b-170">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="fae4b-171">Les fonctions annulables sont généralement utilisées uniquement pour les fonctions asynchrones qui renvoient un seul résultat et doivent gérer l’annulation d’une demande de données.</span><span class="sxs-lookup"><span data-stu-id="fae4b-171">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="fae4b-172">Une fonction ne peut pas être à la fois en continu et annulable.</span><span class="sxs-lookup"><span data-stu-id="fae4b-172">A function cannot be both streaming and cancelable.</span></span> <span data-ttu-id="fae4b-173">Pour plus d’informations, reportez-vous à la remarque à la fin de la [création d’une fonction de diffusion en continu](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="fae4b-173">For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
| `requiresAddress` | <span data-ttu-id="fae4b-174">boolean</span><span class="sxs-lookup"><span data-stu-id="fae4b-174">boolean</span></span>   | <span data-ttu-id="fae4b-175">Non</span><span class="sxs-lookup"><span data-stu-id="fae4b-175">No</span></span> <br/><br/><span data-ttu-id="fae4b-176">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="fae4b-176">Default value is `false`.</span></span> | <span data-ttu-id="fae4b-177">Si `true`votre fonction personnalisée peut accéder à l’adresse de la cellule qui a appelé votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="fae4b-177">If `true`, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="fae4b-178">Pour obtenir l’adresse de la cellule qui a appelé votre fonction personnalisée, utilisez Context. Address dans votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="fae4b-178">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="fae4b-179">Pour plus d’informations, consultez la rubrique relative au [paramètre context de la cellule](/office/dev/add-ins/excel/custom-functions-parameter-options#addressing-cells-context-parameter)Addressing.</span><span class="sxs-lookup"><span data-stu-id="fae4b-179">For more information, see [Addressing cell's context parameter](/office/dev/add-ins/excel/custom-functions-parameter-options#addressing-cells-context-parameter).</span></span> <span data-ttu-id="fae4b-180">Les fonctions personnalisées ne peuvent pas être définies à la fois en diffusion en continu et requiresAddress.</span><span class="sxs-lookup"><span data-stu-id="fae4b-180">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="fae4b-181">Lorsque vous utilisez cette option, le paramètre «invocation» doit être le dernier paramètre passé dans options.</span><span class="sxs-lookup"><span data-stu-id="fae4b-181">When using this option, the 'invocation' parameter must be the last parameter passed in options.</span></span>                                              |
| `stream`          | <span data-ttu-id="fae4b-182">boolean</span><span class="sxs-lookup"><span data-stu-id="fae4b-182">boolean</span></span>   | <span data-ttu-id="fae4b-183">Non</span><span class="sxs-lookup"><span data-stu-id="fae4b-183">No</span></span><br/><br/><span data-ttu-id="fae4b-184">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="fae4b-184">Default value is `false`.</span></span>  | <span data-ttu-id="fae4b-185">Si la valeur est `true`, la fonction peut envoyer une sortie à la cellule à plusieurs reprises, même en cas d’appel unique.</span><span class="sxs-lookup"><span data-stu-id="fae4b-185">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="fae4b-186">Cette option est utile pour des sources de données qui changent rapidement, telles que des valeurs boursières.</span><span class="sxs-lookup"><span data-stu-id="fae4b-186">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="fae4b-187">La fonction ne doit pas utiliser d’instruction `return`.</span><span class="sxs-lookup"><span data-stu-id="fae4b-187">The function should have no `return` statement.</span></span> <span data-ttu-id="fae4b-188">Au lieu de cela, la valeur obtenue est transmise en tant qu’argument de la méthode de rappel `StreamingInvocation.setResult`.</span><span class="sxs-lookup"><span data-stu-id="fae4b-188">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="fae4b-189">Pour plus d’informations, voir [Diffusion en continu de fonctions](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="fae4b-189">For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).</span></span>                                                                                                                                                                |
| `volatile`        | <span data-ttu-id="fae4b-190">boolean</span><span class="sxs-lookup"><span data-stu-id="fae4b-190">boolean</span></span>   | <span data-ttu-id="fae4b-191">Non</span><span class="sxs-lookup"><span data-stu-id="fae4b-191">No</span></span> <br/><br/><span data-ttu-id="fae4b-192">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="fae4b-192">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="fae4b-193">Si la valeur est `true`, la fonction est recalculée à chaque recalcul d’Excel, et plus à chaque fois que les valeurs dépendantes de la formules sont modifiées.</span><span class="sxs-lookup"><span data-stu-id="fae4b-193">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="fae4b-194">Une fonction ne peut pas être à la fois diffusée en continu et volatile.</span><span class="sxs-lookup"><span data-stu-id="fae4b-194">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="fae4b-195">Si les propriétés `stream` et `volatile` sont toutes les deux définies sur `true`, l’option volatile est ignorée.</span><span class="sxs-lookup"><span data-stu-id="fae4b-195">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span>                                                                                                                                                                                                                                                                                             |

### <a name="parameters"></a><span data-ttu-id="fae4b-196">paramètres</span><span class="sxs-lookup"><span data-stu-id="fae4b-196">parameters</span></span>

<span data-ttu-id="fae4b-197">La propriété `parameters` est un tableau d’objets paramètre.</span><span class="sxs-lookup"><span data-stu-id="fae4b-197">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="fae4b-198">Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="fae4b-198">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="fae4b-199">Propriété</span><span class="sxs-lookup"><span data-stu-id="fae4b-199">Property</span></span>  |  <span data-ttu-id="fae4b-200">Type de données</span><span class="sxs-lookup"><span data-stu-id="fae4b-200">Data type</span></span>  |  <span data-ttu-id="fae4b-201">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="fae4b-201">Required</span></span>  |  <span data-ttu-id="fae4b-202">Description</span><span class="sxs-lookup"><span data-stu-id="fae4b-202">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="fae4b-203">string</span><span class="sxs-lookup"><span data-stu-id="fae4b-203">string</span></span>  |  <span data-ttu-id="fae4b-204">Non</span><span class="sxs-lookup"><span data-stu-id="fae4b-204">No</span></span> |  <span data-ttu-id="fae4b-205">Description du paramètre.</span><span class="sxs-lookup"><span data-stu-id="fae4b-205">A description of the parameter.</span></span> <span data-ttu-id="fae4b-206">S’affiche dans intelliSense d’Excel.</span><span class="sxs-lookup"><span data-stu-id="fae4b-206">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="fae4b-207">string</span><span class="sxs-lookup"><span data-stu-id="fae4b-207">string</span></span>  |  <span data-ttu-id="fae4b-208">Non</span><span class="sxs-lookup"><span data-stu-id="fae4b-208">No</span></span>  |  <span data-ttu-id="fae4b-209">Doit être **scalaire** (valeur autre que de tableau) ou **matrice** (tableau bidimensionnel).</span><span class="sxs-lookup"><span data-stu-id="fae4b-209">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="fae4b-210">string</span><span class="sxs-lookup"><span data-stu-id="fae4b-210">string</span></span>  |  <span data-ttu-id="fae4b-211">Oui</span><span class="sxs-lookup"><span data-stu-id="fae4b-211">Yes</span></span>  |  <span data-ttu-id="fae4b-212">Le nom du paramètre.</span><span class="sxs-lookup"><span data-stu-id="fae4b-212">The name of the parameter.</span></span> <span data-ttu-id="fae4b-213">Ce nom s’affiche dans intelliSense d’Excel.</span><span class="sxs-lookup"><span data-stu-id="fae4b-213">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="fae4b-214">string</span><span class="sxs-lookup"><span data-stu-id="fae4b-214">string</span></span>  |  <span data-ttu-id="fae4b-215">Non</span><span class="sxs-lookup"><span data-stu-id="fae4b-215">No</span></span>  |  <span data-ttu-id="fae4b-216">Type de données du paramètre.</span><span class="sxs-lookup"><span data-stu-id="fae4b-216">The data type of the parameter.</span></span> <span data-ttu-id="fae4b-217">Peut être **boolean**, **number**, **string** ou **any** qui vous permet d’utiliser n’importe lequel des trois types précédents.</span><span class="sxs-lookup"><span data-stu-id="fae4b-217">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="fae4b-218">Si cette propriété n’est pas spécifiée, le type de données par défaut est **any**.</span><span class="sxs-lookup"><span data-stu-id="fae4b-218">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="fae4b-219">boolean</span><span class="sxs-lookup"><span data-stu-id="fae4b-219">boolean</span></span> | <span data-ttu-id="fae4b-220">Non</span><span class="sxs-lookup"><span data-stu-id="fae4b-220">No</span></span> | <span data-ttu-id="fae4b-221">Si la valeur est `true`, le paramètre est facultatif.</span><span class="sxs-lookup"><span data-stu-id="fae4b-221">If `true`, the parameter is optional.</span></span> |
|`repeating`| <span data-ttu-id="fae4b-222">boolean</span><span class="sxs-lookup"><span data-stu-id="fae4b-222">boolean</span></span> | <span data-ttu-id="fae4b-223">Non</span><span class="sxs-lookup"><span data-stu-id="fae4b-223">No</span></span> | <span data-ttu-id="fae4b-224">Si `true`, les paramètres sont renseignés à partir d’un tableau spécifié.</span><span class="sxs-lookup"><span data-stu-id="fae4b-224">If `true`, parameters will populate from a specified array.</span></span> <span data-ttu-id="fae4b-225">Notez que les fonctions de tous les paramètres répétitifs sont considérées comme des paramètres facultatifs par définition.</span><span class="sxs-lookup"><span data-stu-id="fae4b-225">Note that functions all repeating parameters are considered optional parameters by definition.</span></span>  |

### <a name="result"></a><span data-ttu-id="fae4b-226">résultat</span><span class="sxs-lookup"><span data-stu-id="fae4b-226">result</span></span>

<span data-ttu-id="fae4b-227">L’objet `result` définit le type des informations renvoyées par la fonction.</span><span class="sxs-lookup"><span data-stu-id="fae4b-227">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="fae4b-228">Le tableau suivant répertorie les propriétés de l’objet `result`.</span><span class="sxs-lookup"><span data-stu-id="fae4b-228">The following table lists the properties of the `result` object.</span></span>

| <span data-ttu-id="fae4b-229">Propriété</span><span class="sxs-lookup"><span data-stu-id="fae4b-229">Property</span></span>         | <span data-ttu-id="fae4b-230">Type de données</span><span class="sxs-lookup"><span data-stu-id="fae4b-230">Data type</span></span> | <span data-ttu-id="fae4b-231">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="fae4b-231">Required</span></span> | <span data-ttu-id="fae4b-232">Description</span><span class="sxs-lookup"><span data-stu-id="fae4b-232">Description</span></span>                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | <span data-ttu-id="fae4b-233">string</span><span class="sxs-lookup"><span data-stu-id="fae4b-233">string</span></span>    | <span data-ttu-id="fae4b-234">Non</span><span class="sxs-lookup"><span data-stu-id="fae4b-234">No</span></span>       | <span data-ttu-id="fae4b-235">Doit être **scalaire** (valeur autre que de tableau) ou **matrice** (tableau bidimensionnel).</span><span class="sxs-lookup"><span data-stu-id="fae4b-235">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="fae4b-236">Mappage des noms de fonction aux métadonnées JSON</span><span class="sxs-lookup"><span data-stu-id="fae4b-236">Associating function names with JSON metadata</span></span>

<span data-ttu-id="fae4b-237">Pour qu’une fonction fonctionne correctement, vous devez associer la propriété de `id` la fonction à l’implémentation JavaScript.</span><span class="sxs-lookup"><span data-stu-id="fae4b-237">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="fae4b-238">Assurez-vous qu’il existe une association, sinon la fonction n’est pas inscrite et n’est pas utilisable dans Excel.</span><span class="sxs-lookup"><span data-stu-id="fae4b-238">Make sure there is an association, otherwise the function will not be registered and not useable in Excel.</span></span> <span data-ttu-id="fae4b-239">L’exemple de code suivant montre comment effectuer l’Association à l' `CustomFunctions.associate()` aide de la méthode.</span><span class="sxs-lookup"><span data-stu-id="fae4b-239">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="fae4b-240">L’exemple définit la fonction personnalisée `add` et associe à l’objet dans le fichier de métadonnées JSON où la valeur de la propriété`id`est**AJOUTER**.</span><span class="sxs-lookup"><span data-stu-id="fae4b-240">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="fae4b-241">Le code JSON suivant illustre les métadonnées JSON associées au code JavaScript de fonction personnalisée précédent.</span><span class="sxs-lookup"><span data-stu-id="fae4b-241">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

```json
{
  "functions": [
    {
      "description": "Add two numbers",
      "id": "ADD",
      "name": "ADD",
      "parameters": [
        {
          "description": "First number",
          "name": "first",
          "type": "number"
        },
        {
          "description": "Second number",
          "name": "second",
          "type": "number"
        }
      ],
      "result": {
        "type": "number"
      }
    }
  ]
}
```

<span data-ttu-id="fae4b-242">N’oubliez pas les meilleures pratiques suivantes lors de la création de fonctions personnalisées dans votre fichier JavaScript et spécifiez les informations correspondantes dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="fae4b-242">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

- <span data-ttu-id="fae4b-243">Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété contient uniquement des points et des caractères alphanumériques.</span><span class="sxs-lookup"><span data-stu-id="fae4b-243">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

- <span data-ttu-id="fae4b-244">Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété est unique dans l’étendue du fichier.</span><span class="sxs-lookup"><span data-stu-id="fae4b-244">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="fae4b-245">Autrement dit, aucun objet fonction dans le fichier de métadonnées ne doit pas avoir la même`id`valeur.</span><span class="sxs-lookup"><span data-stu-id="fae4b-245">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

- <span data-ttu-id="fae4b-246">Ne modifiez pas la valeur d’une`id` propriété dans le fichier de métadonnées JSON après qu’elle ait été mappée à un nom de fonction JavaScript correspondante.</span><span class="sxs-lookup"><span data-stu-id="fae4b-246">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="fae4b-247">Vous pouvez modifier le nom de fonction que voient les utilisateurs finaux dans Excel en mettant à jour la `name` propriété dans le fichier de métadonnées JSON, mais vous ne devez jamais changer la valeur d’une `id` propriété après qu’elle a été établie.</span><span class="sxs-lookup"><span data-stu-id="fae4b-247">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

- <span data-ttu-id="fae4b-248">Dans le fichier JavaScript, spécifiez une association de fonctions `CustomFunctions.associate` personnalisées à l’aide de after each.</span><span class="sxs-lookup"><span data-stu-id="fae4b-248">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="fae4b-249">L’exemple suivant montre les métadonnées JSON correspondant aux fonctions définies dans cet exemple de code JavaScript.</span><span class="sxs-lookup"><span data-stu-id="fae4b-249">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="fae4b-250">Les `id` valeurs `name` de la propriété et sont en majuscules, ce qui est recommandé lors de la description de vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="fae4b-250">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="fae4b-251">Vous n’avez besoin d’ajouter ce JSON que si vous préparez votre propre fichier JSON manuellement et non à l’aide de la génération automatique.</span><span class="sxs-lookup"><span data-stu-id="fae4b-251">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="fae4b-252">Pour plus d’informations sur la génération automatique, voir [Create JSON Metadata for Custom Functions](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="fae4b-252">For more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      ...
    },
    {
      "id": "INCREMENT",
      "name": "INCREMENT",
      ...
    }
  ]
}
```

## <a name="next-steps"></a><span data-ttu-id="fae4b-253">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="fae4b-253">Next steps</span></span>

<span data-ttu-id="fae4b-254">Découvrez les [meilleures pratiques de dénomination de votre fonction](custom-functions-naming.md) ou Découvrez comment [localiser votre fonction](custom-functions-localize.md) à l’aide de la méthode JSON manuscrite décrite précédemment.</span><span class="sxs-lookup"><span data-stu-id="fae4b-254">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="fae4b-255">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="fae4b-255">See also</span></span>

- [<span data-ttu-id="fae4b-256">Générer automatiquement des métadonnées JSON pour des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="fae4b-256">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="fae4b-257">Options des paramètres de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="fae4b-257">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
- [<span data-ttu-id="fae4b-258">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="fae4b-258">Create custom functions in Excel</span></span>](custom-functions-overview.md)
