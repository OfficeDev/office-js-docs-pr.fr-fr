---
ms.date: 11/06/2020
description: Définissez des métadonnées JSON pour les fonctions personnalisées dans Excel et associez vos ID de fonction et propriétés de nom.
title: Créer manuellement des métadonnées JSON pour les fonctions personnalisées dans Excel
localization_priority: Normal
ms.openlocfilehash: adbcbb9d2705a38b1ed9ff5cdffa6162b9d93a9c
ms.sourcegitcommit: 5bfd1e9956485c140179dfcc9d210c4c5a49a789
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/13/2020
ms.locfileid: "49071640"
---
# <a name="manually-create-json-metadata-for-custom-functions"></a><span data-ttu-id="031d7-103">Créer manuellement des métadonnées JSON pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="031d7-103">Manually create JSON metadata for custom functions</span></span>

<span data-ttu-id="031d7-104">Comme décrit dans l’article [vue d’ensemble des fonctions personnalisées](custom-functions-overview.md) , un projet de fonctions personnalisées doit inclure un fichier de métadonnées JSON et un fichier script (JavaScript ou machine à écriture) pour enregistrer une fonction, le rendant ainsi disponible.</span><span class="sxs-lookup"><span data-stu-id="031d7-104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to register a function, making it available for use.</span></span> <span data-ttu-id="031d7-105">Les fonctions personnalisées sont inscrites lorsque l’utilisateur exécute le complément pour la première fois et après qu’il est disponible pour le même utilisateur dans tous les classeurs.</span><span class="sxs-lookup"><span data-stu-id="031d7-105">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="031d7-106">Nous vous recommandons d’utiliser la génération automatique JSON lorsque cela est possible au lieu de créer votre propre fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="031d7-106">We recommend using JSON autogeneration when possible instead of creating your own JSON file.</span></span> <span data-ttu-id="031d7-107">La génération automatique est moins sujette aux erreurs de l’utilisateur et les fichiers générés par la génération de `yo office` modèles automatiques incluent déjà cela.</span><span class="sxs-lookup"><span data-stu-id="031d7-107">Autogeneration is less prone to user error and the `yo office` scaffolded files already include this.</span></span> <span data-ttu-id="031d7-108">Pour plus d’informations sur les balises JSDoc et le processus de génération automatique JSON, voir [génération automatique de métadonnées JSON pour les fonctions personnalisées](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="031d7-108">For more information on JSDoc tags and the JSON autogeneration process, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="031d7-109">Toutefois, vous pouvez créer un projet de fonctions personnalisées à partir de zéro.</span><span class="sxs-lookup"><span data-stu-id="031d7-109">However, you can make a custom functions project from scratch.</span></span> <span data-ttu-id="031d7-110">Ce processus nécessite d’effectuer les opérations suivantes :</span><span class="sxs-lookup"><span data-stu-id="031d7-110">This process requires you to:</span></span>

- <span data-ttu-id="031d7-111">Écrivez votre fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="031d7-111">Write your JSON file.</span></span>
- <span data-ttu-id="031d7-112">Vérifiez que votre fichier manifeste est connecté à votre fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="031d7-112">Check that your manifest file is connected to your JSON file.</span></span>
- <span data-ttu-id="031d7-113">Associez les fonctions `id` et les `name` Propriétés dans le fichier de script pour enregistrer vos fonctions.</span><span class="sxs-lookup"><span data-stu-id="031d7-113">Associate your functions' `id` and `name` properties in the script file in order to register your functions.</span></span>

<span data-ttu-id="031d7-114">L’image suivante explique les différences entre l’utilisation `yo office` de fichiers de structure et l’écriture de JSON à partir de zéro.</span><span class="sxs-lookup"><span data-stu-id="031d7-114">The following image explains the differences between using `yo office` scaffold files and writing JSON from scratch.</span></span>

![Image des différences entre l’utilisation de yo Office et l’écriture de votre propre JSON](../images/custom-functions-json.png)

> [!NOTE]
> <span data-ttu-id="031d7-116">N’oubliez pas de connecter votre manifeste au fichier JSON que vous créez, via la `<Resources>` section de votre fichier manifeste XML si vous n’utilisez pas le `yo office` Générateur.</span><span class="sxs-lookup"><span data-stu-id="031d7-116">Remember to connect your manifest to the JSON file you create, through the `<Resources>` section in your XML manifest file if you do not use the `yo office` generator.</span></span>

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a><span data-ttu-id="031d7-117">Création de métadonnées et connexion au manifeste</span><span class="sxs-lookup"><span data-stu-id="031d7-117">Authoring metadata and connecting to the manifest</span></span>

<span data-ttu-id="031d7-118">Créez un fichier JSON dans votre projet et fournissez-y tous les détails sur vos fonctions, telles que les paramètres de la fonction.</span><span class="sxs-lookup"><span data-stu-id="031d7-118">Create a JSON file in your project and provide all the details about your functions in it, such as the function's parameters.</span></span> <span data-ttu-id="031d7-119">Consultez l' [exemple de métadonnées suivant](#json-metadata-example) et [la référence de métadonnées](#metadata-reference) pour obtenir la liste complète des propriétés de fonction.</span><span class="sxs-lookup"><span data-stu-id="031d7-119">See the [following metadata example](#json-metadata-example) and [the metadata reference](#metadata-reference) for a complete list of function properties.</span></span>

<span data-ttu-id="031d7-120">Assurez-vous que votre fichier manifeste XML fait référence à votre fichier JSON dans la `<Resources>` section, comme dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="031d7-120">Ensure your XML manifest file references your JSON file in the `<Resources>` section, similar to the following example.</span></span>

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

## <a name="json-metadata-example"></a><span data-ttu-id="031d7-121">Exemple de métadonnées JSON</span><span class="sxs-lookup"><span data-stu-id="031d7-121">JSON metadata example</span></span>

<span data-ttu-id="031d7-122">L’exemple suivant montre le contenu d’un fichier de métadonnées JSON pour un complément qui définit des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="031d7-122">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="031d7-123">Les sections qui suivent cet exemple fournissent des informations détaillées sur les propriétés individuelles au sein de cet exemple JSON.</span><span class="sxs-lookup"><span data-stu-id="031d7-123">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="031d7-124">Un exemple de fichier JSON complet est disponible dans l’historique de validation du référentiel [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) github.</span><span class="sxs-lookup"><span data-stu-id="031d7-124">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="031d7-125">Lorsque le projet a été ajusté pour générer automatiquement JSON, un échantillon complet de JSON manuscrit est uniquement disponible dans les versions précédentes du projet.</span><span class="sxs-lookup"><span data-stu-id="031d7-125">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="metadata-reference"></a><span data-ttu-id="031d7-126">Référence de métadonnées</span><span class="sxs-lookup"><span data-stu-id="031d7-126">Metadata reference</span></span>

### <a name="functions"></a><span data-ttu-id="031d7-127">fonctions</span><span class="sxs-lookup"><span data-stu-id="031d7-127">functions</span></span>

<span data-ttu-id="031d7-128">La propriété `functions` est un tableau d’objets de fonction personnalisés.</span><span class="sxs-lookup"><span data-stu-id="031d7-128">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="031d7-129">Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="031d7-129">The following table lists the properties of each object.</span></span>

| <span data-ttu-id="031d7-130">Propriété</span><span class="sxs-lookup"><span data-stu-id="031d7-130">Property</span></span>      | <span data-ttu-id="031d7-131">Type de données</span><span class="sxs-lookup"><span data-stu-id="031d7-131">Data type</span></span> | <span data-ttu-id="031d7-132">Requis</span><span class="sxs-lookup"><span data-stu-id="031d7-132">Required</span></span> | <span data-ttu-id="031d7-133">Description</span><span class="sxs-lookup"><span data-stu-id="031d7-133">Description</span></span>                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | <span data-ttu-id="031d7-134">string</span><span class="sxs-lookup"><span data-stu-id="031d7-134">string</span></span>    | <span data-ttu-id="031d7-135">Non</span><span class="sxs-lookup"><span data-stu-id="031d7-135">No</span></span>       | <span data-ttu-id="031d7-136">Description de la fonction que voient les utilisateurs finaux dans Excel.</span><span class="sxs-lookup"><span data-stu-id="031d7-136">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="031d7-137">Par exemple, **convertit une valeur Celsius en valeur Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="031d7-137">For example, **Converts a Celsius value to Fahrenheit**.</span></span>                                                            |
| `helpUrl`     | <span data-ttu-id="031d7-138">string</span><span class="sxs-lookup"><span data-stu-id="031d7-138">string</span></span>    | <span data-ttu-id="031d7-139">Non</span><span class="sxs-lookup"><span data-stu-id="031d7-139">No</span></span>       | <span data-ttu-id="031d7-140">URL fournissant des informations sur la fonction</span><span class="sxs-lookup"><span data-stu-id="031d7-140">URL that provides information about the function.</span></span> <span data-ttu-id="031d7-141">(elle est affichée dans un volet des tâches). Par exemple, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span><span class="sxs-lookup"><span data-stu-id="031d7-141">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span>                      |
| `id`          | <span data-ttu-id="031d7-142">string</span><span class="sxs-lookup"><span data-stu-id="031d7-142">string</span></span>    | <span data-ttu-id="031d7-143">Oui</span><span class="sxs-lookup"><span data-stu-id="031d7-143">Yes</span></span>      | <span data-ttu-id="031d7-144">Un ID unique pour la fonction.</span><span class="sxs-lookup"><span data-stu-id="031d7-144">A unique ID for the function.</span></span> <span data-ttu-id="031d7-145">Cet ID peut contenir uniquement des points et caractères alphanumériques et ne doit pas être modifié une fois défini.</span><span class="sxs-lookup"><span data-stu-id="031d7-145">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span>                                            |
| `name`        | <span data-ttu-id="031d7-146">string</span><span class="sxs-lookup"><span data-stu-id="031d7-146">string</span></span>    | <span data-ttu-id="031d7-147">Oui</span><span class="sxs-lookup"><span data-stu-id="031d7-147">Yes</span></span>      | <span data-ttu-id="031d7-148">Nom de la fonction que voient les utilisateurs finaux dans Excel.</span><span class="sxs-lookup"><span data-stu-id="031d7-148">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="031d7-149">Dans Excel, le nom de cette fonction est préfixé par l’espace de noms des fonctions personnalisées qui est spécifié dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="031d7-149">In Excel, this function name is prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `options`     | <span data-ttu-id="031d7-150">object</span><span class="sxs-lookup"><span data-stu-id="031d7-150">object</span></span>    | <span data-ttu-id="031d7-151">Non</span><span class="sxs-lookup"><span data-stu-id="031d7-151">No</span></span>       | <span data-ttu-id="031d7-152">Vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction.</span><span class="sxs-lookup"><span data-stu-id="031d7-152">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="031d7-153">Reportez-vous aux [options](#options) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="031d7-153">See [options](#options) for details.</span></span>                                                          |
| `parameters`  | <span data-ttu-id="031d7-154">tableau</span><span class="sxs-lookup"><span data-stu-id="031d7-154">array</span></span>     | <span data-ttu-id="031d7-155">Oui</span><span class="sxs-lookup"><span data-stu-id="031d7-155">Yes</span></span>      | <span data-ttu-id="031d7-156">Tableau qui définit les paramètres d’entrée de la fonction.</span><span class="sxs-lookup"><span data-stu-id="031d7-156">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="031d7-157">Pour plus d’informations, consultez la rubrique [paramètres](#parameters) .</span><span class="sxs-lookup"><span data-stu-id="031d7-157">See [parameters](#parameters) for details.</span></span>                                                                             |
| `result`      | <span data-ttu-id="031d7-158">objet</span><span class="sxs-lookup"><span data-stu-id="031d7-158">object</span></span>    | <span data-ttu-id="031d7-159">Oui</span><span class="sxs-lookup"><span data-stu-id="031d7-159">Yes</span></span>      | <span data-ttu-id="031d7-160">Objet qui définit le type d’informations renvoyées par la fonction.</span><span class="sxs-lookup"><span data-stu-id="031d7-160">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="031d7-161">Reportez-vous au [résultat](#result) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="031d7-161">See [result](#result) for details.</span></span>                                                                 |

### <a name="options"></a><span data-ttu-id="031d7-162">options</span><span class="sxs-lookup"><span data-stu-id="031d7-162">options</span></span>

<span data-ttu-id="031d7-163">L’objet `options` vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction.</span><span class="sxs-lookup"><span data-stu-id="031d7-163">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="031d7-164">Le tableau suivant répertorie les propriétés de l’objet `options`.</span><span class="sxs-lookup"><span data-stu-id="031d7-164">The following table lists the properties of the `options` object.</span></span>

| <span data-ttu-id="031d7-165">Propriété</span><span class="sxs-lookup"><span data-stu-id="031d7-165">Property</span></span>          | <span data-ttu-id="031d7-166">Type de données</span><span class="sxs-lookup"><span data-stu-id="031d7-166">Data type</span></span> | <span data-ttu-id="031d7-167">Requis</span><span class="sxs-lookup"><span data-stu-id="031d7-167">Required</span></span>                               | <span data-ttu-id="031d7-168">Description</span><span class="sxs-lookup"><span data-stu-id="031d7-168">Description</span></span> |
| :---------------- | :-------- | :------------------------------------- | :---------- |
| `cancelable`      | <span data-ttu-id="031d7-169">boolean</span><span class="sxs-lookup"><span data-stu-id="031d7-169">boolean</span></span>   | <span data-ttu-id="031d7-170">Non</span><span class="sxs-lookup"><span data-stu-id="031d7-170">No</span></span><br/><br/><span data-ttu-id="031d7-171">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="031d7-171">Default value is `false`.</span></span>  | <span data-ttu-id="031d7-172">Si la valeur est `true`, Excel appelle le gestionnaire `CancelableInvocation` chaque fois que l’utilisateur effectue une action ayant pour effet d’annuler la fonction, par exemple, en déclenchant manuellement un recalcul ou en modifiant une cellule référencée par la fonction.</span><span class="sxs-lookup"><span data-stu-id="031d7-172">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="031d7-173">Les fonctions annulables sont généralement utilisées uniquement pour les fonctions asynchrones qui renvoient un seul résultat et doivent gérer l’annulation d’une demande de données.</span><span class="sxs-lookup"><span data-stu-id="031d7-173">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="031d7-174">Une fonction ne peut pas être à la fois en continu et annulable.</span><span class="sxs-lookup"><span data-stu-id="031d7-174">A function cannot be both streaming and cancelable.</span></span> <span data-ttu-id="031d7-175">Pour plus d’informations, reportez-vous à la remarque à la fin de la [création d’une fonction de diffusion en continu](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="031d7-175">For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
| `requiresAddress` | <span data-ttu-id="031d7-176">boolean</span><span class="sxs-lookup"><span data-stu-id="031d7-176">boolean</span></span>   | <span data-ttu-id="031d7-177">Non</span><span class="sxs-lookup"><span data-stu-id="031d7-177">No</span></span> <br/><br/><span data-ttu-id="031d7-178">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="031d7-178">Default value is `false`.</span></span> | <span data-ttu-id="031d7-179">Si `true` votre fonction personnalisée peut accéder à l’adresse de la cellule qui a appelé votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="031d7-179">If `true`, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="031d7-180">Pour obtenir l’adresse de la cellule qui a appelé votre fonction personnalisée, utilisez Context. Address dans votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="031d7-180">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="031d7-181">Les fonctions personnalisées ne peuvent pas être définies à la fois en diffusion en continu et requiresAddress.</span><span class="sxs-lookup"><span data-stu-id="031d7-181">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="031d7-182">Lorsque vous utilisez cette option, le paramètre « invocation » doit être le dernier paramètre passé dans options.</span><span class="sxs-lookup"><span data-stu-id="031d7-182">When using this option, the 'invocation' parameter must be the last parameter passed in options.</span></span> |
| `stream`          | <span data-ttu-id="031d7-183">boolean</span><span class="sxs-lookup"><span data-stu-id="031d7-183">boolean</span></span>   | <span data-ttu-id="031d7-184">Non</span><span class="sxs-lookup"><span data-stu-id="031d7-184">No</span></span><br/><br/><span data-ttu-id="031d7-185">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="031d7-185">Default value is `false`.</span></span>  | <span data-ttu-id="031d7-186">Si la valeur est `true`, la fonction peut envoyer une sortie à la cellule à plusieurs reprises, même en cas d’appel unique.</span><span class="sxs-lookup"><span data-stu-id="031d7-186">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="031d7-187">Cette option est utile pour des sources de données qui changent rapidement, telles que des valeurs boursières.</span><span class="sxs-lookup"><span data-stu-id="031d7-187">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="031d7-188">La fonction ne doit pas utiliser d’instruction `return`.</span><span class="sxs-lookup"><span data-stu-id="031d7-188">The function should have no `return` statement.</span></span> <span data-ttu-id="031d7-189">Au lieu de cela, la valeur obtenue est transmise en tant qu’argument de la méthode de rappel `StreamingInvocation.setResult`.</span><span class="sxs-lookup"><span data-stu-id="031d7-189">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="031d7-190">Pour plus d’informations, voir [Diffusion en continu de fonctions](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="031d7-190">For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
| `volatile`        | <span data-ttu-id="031d7-191">boolean</span><span class="sxs-lookup"><span data-stu-id="031d7-191">boolean</span></span>   | <span data-ttu-id="031d7-192">Non</span><span class="sxs-lookup"><span data-stu-id="031d7-192">No</span></span> <br/><br/><span data-ttu-id="031d7-193">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="031d7-193">Default value is `false`.</span></span> | <span data-ttu-id="031d7-194">Si `true` , la fonction recalcule chaque fois qu’Excel recalcule, et non uniquement lorsque les valeurs dépendantes de la formule ont été modifiées.</span><span class="sxs-lookup"><span data-stu-id="031d7-194">If `true`, the function recalculates each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="031d7-195">Une fonction ne peut pas être à la fois diffusée en continu et volatile.</span><span class="sxs-lookup"><span data-stu-id="031d7-195">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="031d7-196">Si les propriétés `stream` et `volatile` sont toutes les deux définies sur `true`, l’option volatile est ignorée.</span><span class="sxs-lookup"><span data-stu-id="031d7-196">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

### <a name="parameters"></a><span data-ttu-id="031d7-197">paramètres</span><span class="sxs-lookup"><span data-stu-id="031d7-197">parameters</span></span>

<span data-ttu-id="031d7-198">La propriété `parameters` est un tableau d’objets paramètre.</span><span class="sxs-lookup"><span data-stu-id="031d7-198">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="031d7-199">Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="031d7-199">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="031d7-200">Propriété</span><span class="sxs-lookup"><span data-stu-id="031d7-200">Property</span></span>  |  <span data-ttu-id="031d7-201">Type de données</span><span class="sxs-lookup"><span data-stu-id="031d7-201">Data type</span></span>  |  <span data-ttu-id="031d7-202">Requis</span><span class="sxs-lookup"><span data-stu-id="031d7-202">Required</span></span>  |  <span data-ttu-id="031d7-203">Description</span><span class="sxs-lookup"><span data-stu-id="031d7-203">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="031d7-204">string</span><span class="sxs-lookup"><span data-stu-id="031d7-204">string</span></span>  |  <span data-ttu-id="031d7-205">Non</span><span class="sxs-lookup"><span data-stu-id="031d7-205">No</span></span> |  <span data-ttu-id="031d7-206">Description du paramètre.</span><span class="sxs-lookup"><span data-stu-id="031d7-206">A description of the parameter.</span></span> <span data-ttu-id="031d7-207">Elle s’affiche dans IntelliSense d’Excel.</span><span class="sxs-lookup"><span data-stu-id="031d7-207">This is displayed in Excel's IntelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="031d7-208">string</span><span class="sxs-lookup"><span data-stu-id="031d7-208">string</span></span>  |  <span data-ttu-id="031d7-209">Non</span><span class="sxs-lookup"><span data-stu-id="031d7-209">No</span></span>  |  <span data-ttu-id="031d7-210">Doit être **scalaire** (valeur autre que de tableau) ou **matrice** (tableau bidimensionnel).</span><span class="sxs-lookup"><span data-stu-id="031d7-210">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="031d7-211">string</span><span class="sxs-lookup"><span data-stu-id="031d7-211">string</span></span>  |  <span data-ttu-id="031d7-212">Oui</span><span class="sxs-lookup"><span data-stu-id="031d7-212">Yes</span></span>  |  <span data-ttu-id="031d7-213">Le nom du paramètre.</span><span class="sxs-lookup"><span data-stu-id="031d7-213">The name of the parameter.</span></span> <span data-ttu-id="031d7-214">Ce nom s’affiche dans IntelliSense d’Excel.</span><span class="sxs-lookup"><span data-stu-id="031d7-214">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="031d7-215">string</span><span class="sxs-lookup"><span data-stu-id="031d7-215">string</span></span>  |  <span data-ttu-id="031d7-216">Non</span><span class="sxs-lookup"><span data-stu-id="031d7-216">No</span></span>  |  <span data-ttu-id="031d7-217">Type de données du paramètre.</span><span class="sxs-lookup"><span data-stu-id="031d7-217">The data type of the parameter.</span></span> <span data-ttu-id="031d7-218">Peut être **boolean** , **number** , **string** ou **any** qui vous permet d’utiliser n’importe lequel des trois types précédents.</span><span class="sxs-lookup"><span data-stu-id="031d7-218">Can be **boolean** , **number** , **string** , or **any** , which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="031d7-219">Si cette propriété n’est pas spécifiée, le type de données par défaut est **any**.</span><span class="sxs-lookup"><span data-stu-id="031d7-219">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="031d7-220">boolean</span><span class="sxs-lookup"><span data-stu-id="031d7-220">boolean</span></span> | <span data-ttu-id="031d7-221">Non</span><span class="sxs-lookup"><span data-stu-id="031d7-221">No</span></span> | <span data-ttu-id="031d7-222">Si la valeur est `true`, le paramètre est facultatif.</span><span class="sxs-lookup"><span data-stu-id="031d7-222">If `true`, the parameter is optional.</span></span> |
|`repeating`| <span data-ttu-id="031d7-223">boolean</span><span class="sxs-lookup"><span data-stu-id="031d7-223">boolean</span></span> | <span data-ttu-id="031d7-224">Non</span><span class="sxs-lookup"><span data-stu-id="031d7-224">No</span></span> | <span data-ttu-id="031d7-225">Si `true` , les paramètres sont renseignés à partir d’un tableau spécifié.</span><span class="sxs-lookup"><span data-stu-id="031d7-225">If `true`, parameters populate from a specified array.</span></span> <span data-ttu-id="031d7-226">Notez que les fonctions de tous les paramètres répétitifs sont considérées comme des paramètres facultatifs par définition.</span><span class="sxs-lookup"><span data-stu-id="031d7-226">Note that functions all repeating parameters are considered optional parameters by definition.</span></span>  |

### <a name="result"></a><span data-ttu-id="031d7-227">résultat</span><span class="sxs-lookup"><span data-stu-id="031d7-227">result</span></span>

<span data-ttu-id="031d7-228">L’objet `result` définit le type des informations renvoyées par la fonction.</span><span class="sxs-lookup"><span data-stu-id="031d7-228">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="031d7-229">Le tableau suivant répertorie les propriétés de l’objet `result`.</span><span class="sxs-lookup"><span data-stu-id="031d7-229">The following table lists the properties of the `result` object.</span></span>

| <span data-ttu-id="031d7-230">Propriété</span><span class="sxs-lookup"><span data-stu-id="031d7-230">Property</span></span>         | <span data-ttu-id="031d7-231">Type de données</span><span class="sxs-lookup"><span data-stu-id="031d7-231">Data type</span></span> | <span data-ttu-id="031d7-232">Requis</span><span class="sxs-lookup"><span data-stu-id="031d7-232">Required</span></span> | <span data-ttu-id="031d7-233">Description</span><span class="sxs-lookup"><span data-stu-id="031d7-233">Description</span></span>                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | <span data-ttu-id="031d7-234">string</span><span class="sxs-lookup"><span data-stu-id="031d7-234">string</span></span>    | <span data-ttu-id="031d7-235">Non</span><span class="sxs-lookup"><span data-stu-id="031d7-235">No</span></span>       | <span data-ttu-id="031d7-236">Doit être **scalaire** (valeur autre que de tableau) ou **matrice** (tableau bidimensionnel).</span><span class="sxs-lookup"><span data-stu-id="031d7-236">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="031d7-237">Mappage des noms de fonction aux métadonnées JSON</span><span class="sxs-lookup"><span data-stu-id="031d7-237">Associating function names with JSON metadata</span></span>

<span data-ttu-id="031d7-238">Pour qu’une fonction fonctionne correctement, vous devez associer la propriété de la fonction `id` à l’implémentation JavaScript.</span><span class="sxs-lookup"><span data-stu-id="031d7-238">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="031d7-239">Assurez-vous qu’il existe une association, sinon la fonction ne sera pas enregistrée et n’est pas utilisable dans Excel.</span><span class="sxs-lookup"><span data-stu-id="031d7-239">Make sure there is an association, otherwise the function won't be registered and isn't useable in Excel.</span></span> <span data-ttu-id="031d7-240">L’exemple de code suivant montre comment effectuer l’Association à l’aide de la `CustomFunctions.associate()` méthode.</span><span class="sxs-lookup"><span data-stu-id="031d7-240">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="031d7-241">L’exemple définit la fonction personnalisée `add` et associe à l’objet dans le fichier de métadonnées JSON où la valeur de la propriété`id`est **AJOUTER**.</span><span class="sxs-lookup"><span data-stu-id="031d7-241">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

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

<span data-ttu-id="031d7-242">Le code JSON suivant illustre les métadonnées JSON associées au code JavaScript de fonction personnalisée précédent.</span><span class="sxs-lookup"><span data-stu-id="031d7-242">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

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

<span data-ttu-id="031d7-243">N’oubliez pas les meilleures pratiques suivantes lors de la création de fonctions personnalisées dans votre fichier JavaScript et spécifiez les informations correspondantes dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="031d7-243">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

- <span data-ttu-id="031d7-244">Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété contient uniquement des points et des caractères alphanumériques.</span><span class="sxs-lookup"><span data-stu-id="031d7-244">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

- <span data-ttu-id="031d7-245">Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété est unique dans l’étendue du fichier.</span><span class="sxs-lookup"><span data-stu-id="031d7-245">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="031d7-246">Autrement dit, aucun objet fonction dans le fichier de métadonnées ne doit pas avoir la même`id`valeur.</span><span class="sxs-lookup"><span data-stu-id="031d7-246">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

- <span data-ttu-id="031d7-247">Ne modifiez pas la valeur d’une`id` propriété dans le fichier de métadonnées JSON après qu’elle ait été mappée à un nom de fonction JavaScript correspondante.</span><span class="sxs-lookup"><span data-stu-id="031d7-247">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="031d7-248">Vous pouvez modifier le nom de fonction que voient les utilisateurs finaux dans Excel en mettant à jour la `name` propriété dans le fichier de métadonnées JSON, mais vous ne devez jamais changer la valeur d’une `id` propriété après qu’elle a été établie.</span><span class="sxs-lookup"><span data-stu-id="031d7-248">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

- <span data-ttu-id="031d7-249">Dans le fichier JavaScript, spécifiez une association de fonctions personnalisées à l’aide de `CustomFunctions.associate` after each.</span><span class="sxs-lookup"><span data-stu-id="031d7-249">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="031d7-250">L’exemple suivant montre les métadonnées JSON qui correspondent aux fonctions définies dans l’exemple de code JavaScript précédent.</span><span class="sxs-lookup"><span data-stu-id="031d7-250">The following sample shows the JSON metadata that corresponds to the functions defined in the preceding JavaScript code sample.</span></span> <span data-ttu-id="031d7-251">Les `id` valeurs de la `name` propriété et sont en majuscules, ce qui est recommandé lors de la description de vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="031d7-251">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="031d7-252">Vous n’avez besoin d’ajouter ce JSON que si vous préparez votre propre fichier JSON manuellement et non à l’aide de la génération automatique.</span><span class="sxs-lookup"><span data-stu-id="031d7-252">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="031d7-253">Pour plus d’informations sur la génération automatique, voir génération automatique [de métadonnées JSON pour les fonctions personnalisées](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="031d7-253">For more information on autogeneration, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="031d7-254">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="031d7-254">Next steps</span></span>

<span data-ttu-id="031d7-255">Découvrez les [meilleures pratiques de dénomination de votre fonction](custom-functions-naming.md) ou Découvrez comment [localiser votre fonction](custom-functions-localize.md) à l’aide de la méthode JSON manuscrite décrite précédemment.</span><span class="sxs-lookup"><span data-stu-id="031d7-255">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="031d7-256">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="031d7-256">See also</span></span>

- [<span data-ttu-id="031d7-257">Générer automatiquement des métadonnées JSON pour des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="031d7-257">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="031d7-258">Options des paramètres de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="031d7-258">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
- [<span data-ttu-id="031d7-259">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="031d7-259">Create custom functions in Excel</span></span>](custom-functions-overview.md)
