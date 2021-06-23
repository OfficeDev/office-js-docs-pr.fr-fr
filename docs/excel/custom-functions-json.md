---
ms.date: 12/22/2020
description: Définissez les métadonnées JSON pour les fonctions personnalisées Excel et associez votre ID de fonction et vos propriétés de nom.
title: Créer manuellement des métadonnées JSON pour les fonctions personnalisées dans Excel
localization_priority: Normal
ms.openlocfilehash: 514eacba5045d160eb6f3d4823adbd8c2f45292a
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075900"
---
# <a name="manually-create-json-metadata-for-custom-functions"></a><span data-ttu-id="a0bab-103">Créer manuellement des métadonnées JSON pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="a0bab-103">Manually create JSON metadata for custom functions</span></span>

<span data-ttu-id="a0bab-104">Comme décrit dans l’article de vue d’ensemble des fonctions [personnalisées,](custom-functions-overview.md) un projet de fonctions personnalisées doit inclure à la fois un fichier de métadonnées JSON et un fichier de script (JavaScript ou TypeScript) pour inscrire une fonction, ce qui le rend disponible pour utilisation.</span><span class="sxs-lookup"><span data-stu-id="a0bab-104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to register a function, making it available for use.</span></span> <span data-ttu-id="a0bab-105">Les fonctions personnalisées sont enregistrées lorsque l’utilisateur exécute le add-in pour la première fois et après cela sont disponibles pour le même utilisateur dans tous les workbooks.</span><span class="sxs-lookup"><span data-stu-id="a0bab-105">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="a0bab-106">Nous vous recommandons d’utiliser la génération automatique JSON lorsque cela est possible au lieu de créer votre propre fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="a0bab-106">We recommend using JSON autogeneration when possible instead of creating your own JSON file.</span></span> <span data-ttu-id="a0bab-107">La génération automatique est moins sujette aux erreurs de l’utilisateur et les fichiers `yo office` échafaudés l’incluent déjà.</span><span class="sxs-lookup"><span data-stu-id="a0bab-107">Autogeneration is less prone to user error and the `yo office` scaffolded files already include this.</span></span> <span data-ttu-id="a0bab-108">Pour plus d’informations sur les balises JSDoc et le processus de génération automatique JSON, voir métadonnées JSON de génération automatique [pour les fonctions personnalisées](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="a0bab-108">For more information on JSDoc tags and the JSON autogeneration process, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="a0bab-109">Toutefois, vous pouvez créer un projet de fonctions personnalisées à partir de zéro.</span><span class="sxs-lookup"><span data-stu-id="a0bab-109">However, you can make a custom functions project from scratch.</span></span> <span data-ttu-id="a0bab-110">Ce processus nécessite que vous :</span><span class="sxs-lookup"><span data-stu-id="a0bab-110">This process requires you to:</span></span>

- <span data-ttu-id="a0bab-111">Écrivez votre fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="a0bab-111">Write your JSON file.</span></span>
- <span data-ttu-id="a0bab-112">Vérifiez que votre fichier manifeste est connecté à votre fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="a0bab-112">Check that your manifest file is connected to your JSON file.</span></span>
- <span data-ttu-id="a0bab-113">Associez les propriétés et les fonctions de vos fonctions dans `id` le fichier de script afin `name` d’inscrire vos fonctions.</span><span class="sxs-lookup"><span data-stu-id="a0bab-113">Associate your functions' `id` and `name` properties in the script file in order to register your functions.</span></span>

<span data-ttu-id="a0bab-114">L’image suivante explique les différences entre l’utilisation de fichiers de la `yo office` échafaudage et l’écriture de JSON à partir de zéro.</span><span class="sxs-lookup"><span data-stu-id="a0bab-114">The following image explains the differences between using `yo office` scaffold files and writing JSON from scratch.</span></span>

![Image des différences entre l’utilisation de Yo Office et l’écriture de votre propre JSON.](../images/custom-functions-json.png)

> [!NOTE]
> <span data-ttu-id="a0bab-116">N’oubliez pas de connecter votre manifeste au fichier JSON que vous créez, via la section de votre fichier manifeste XML si vous `<Resources>` n’utilisez pas le `yo office` générateur.</span><span class="sxs-lookup"><span data-stu-id="a0bab-116">Remember to connect your manifest to the JSON file you create, through the `<Resources>` section in your XML manifest file if you do not use the `yo office` generator.</span></span>

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a><span data-ttu-id="a0bab-117">Créer des métadonnées et se connecter au manifeste</span><span class="sxs-lookup"><span data-stu-id="a0bab-117">Authoring metadata and connecting to the manifest</span></span>

<span data-ttu-id="a0bab-118">Créez un fichier JSON dans votre projet et fournissez tous les détails sur vos fonctions, telles que les paramètres de la fonction.</span><span class="sxs-lookup"><span data-stu-id="a0bab-118">Create a JSON file in your project and provide all the details about your functions in it, such as the function's parameters.</span></span> <span data-ttu-id="a0bab-119">Consultez [l’exemple de métadonnées suivant](#json-metadata-example) [et la référence des métadonnées](#metadata-reference) pour obtenir la liste complète des propriétés de la fonction.</span><span class="sxs-lookup"><span data-stu-id="a0bab-119">See the [following metadata example](#json-metadata-example) and [the metadata reference](#metadata-reference) for a complete list of function properties.</span></span>

<span data-ttu-id="a0bab-120">Assurez-vous que votre fichier manifeste XML fait référence à votre fichier JSON dans la `<Resources>` section, comme dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="a0bab-120">Ensure your XML manifest file references your JSON file in the `<Resources>` section, similar to the following example.</span></span>

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

## <a name="json-metadata-example"></a><span data-ttu-id="a0bab-121">Exemple de métadonnées JSON</span><span class="sxs-lookup"><span data-stu-id="a0bab-121">JSON metadata example</span></span>

<span data-ttu-id="a0bab-122">L’exemple suivant montre le contenu d’un fichier de métadonnées JSON pour un complément qui définit des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="a0bab-122">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="a0bab-123">Les sections qui suivent cet exemple fournissent des informations détaillées sur les propriétés individuelles au sein de cet exemple JSON.</span><span class="sxs-lookup"><span data-stu-id="a0bab-123">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="a0bab-124">Un exemple complet de fichier JSON est disponible dans [officeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub’historique de validation du référentiel.</span><span class="sxs-lookup"><span data-stu-id="a0bab-124">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="a0bab-125">Comme le projet a été ajusté pour générer automatiquement JSON, un échantillon complet de JSON manuscrit n’est disponible que dans les versions précédentes du projet.</span><span class="sxs-lookup"><span data-stu-id="a0bab-125">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="metadata-reference"></a><span data-ttu-id="a0bab-126">Référence des métadonnées</span><span class="sxs-lookup"><span data-stu-id="a0bab-126">Metadata reference</span></span>

### <a name="functions"></a><span data-ttu-id="a0bab-127">fonctions</span><span class="sxs-lookup"><span data-stu-id="a0bab-127">functions</span></span>

<span data-ttu-id="a0bab-128">La propriété `functions` est un tableau d’objets de fonction personnalisés.</span><span class="sxs-lookup"><span data-stu-id="a0bab-128">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="a0bab-129">Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="a0bab-129">The following table lists the properties of each object.</span></span>

| <span data-ttu-id="a0bab-130">Propriété</span><span class="sxs-lookup"><span data-stu-id="a0bab-130">Property</span></span>      | <span data-ttu-id="a0bab-131">Type de données</span><span class="sxs-lookup"><span data-stu-id="a0bab-131">Data type</span></span> | <span data-ttu-id="a0bab-132">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="a0bab-132">Required</span></span> | <span data-ttu-id="a0bab-133">Description</span><span class="sxs-lookup"><span data-stu-id="a0bab-133">Description</span></span>                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | <span data-ttu-id="a0bab-134">string</span><span class="sxs-lookup"><span data-stu-id="a0bab-134">string</span></span>    | <span data-ttu-id="a0bab-135">Non</span><span class="sxs-lookup"><span data-stu-id="a0bab-135">No</span></span>       | <span data-ttu-id="a0bab-136">Description de la fonction que voient les utilisateurs finaux dans Excel.</span><span class="sxs-lookup"><span data-stu-id="a0bab-136">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="a0bab-137">Par exemple, **convertit une valeur Celsius en valeur Fahrenheit**.</span><span class="sxs-lookup"><span data-stu-id="a0bab-137">For example, **Converts a Celsius value to Fahrenheit**.</span></span>                                                            |
| `helpUrl`     | <span data-ttu-id="a0bab-138">string</span><span class="sxs-lookup"><span data-stu-id="a0bab-138">string</span></span>    | <span data-ttu-id="a0bab-139">Non</span><span class="sxs-lookup"><span data-stu-id="a0bab-139">No</span></span>       | <span data-ttu-id="a0bab-140">URL fournissant des informations sur la fonction</span><span class="sxs-lookup"><span data-stu-id="a0bab-140">URL that provides information about the function.</span></span> <span data-ttu-id="a0bab-141">(elle est affichée dans un volet des tâches). Par exemple, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span><span class="sxs-lookup"><span data-stu-id="a0bab-141">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span>                      |
| `id`          | <span data-ttu-id="a0bab-142">string</span><span class="sxs-lookup"><span data-stu-id="a0bab-142">string</span></span>    | <span data-ttu-id="a0bab-143">Oui</span><span class="sxs-lookup"><span data-stu-id="a0bab-143">Yes</span></span>      | <span data-ttu-id="a0bab-144">Un ID unique pour la fonction.</span><span class="sxs-lookup"><span data-stu-id="a0bab-144">A unique ID for the function.</span></span> <span data-ttu-id="a0bab-145">Cet ID peut contenir uniquement des points et caractères alphanumériques et ne doit pas être modifié une fois défini.</span><span class="sxs-lookup"><span data-stu-id="a0bab-145">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span>                                            |
| `name`        | <span data-ttu-id="a0bab-146">string</span><span class="sxs-lookup"><span data-stu-id="a0bab-146">string</span></span>    | <span data-ttu-id="a0bab-147">Oui</span><span class="sxs-lookup"><span data-stu-id="a0bab-147">Yes</span></span>      | <span data-ttu-id="a0bab-148">Nom de la fonction que voient les utilisateurs finaux dans Excel.</span><span class="sxs-lookup"><span data-stu-id="a0bab-148">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="a0bab-149">Dans Excel, le nom de cette fonction est précédé de l’espace de noms des fonctions personnalisées spécifié dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="a0bab-149">In Excel, this function name is prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `options`     | <span data-ttu-id="a0bab-150">object</span><span class="sxs-lookup"><span data-stu-id="a0bab-150">object</span></span>    | <span data-ttu-id="a0bab-151">Non</span><span class="sxs-lookup"><span data-stu-id="a0bab-151">No</span></span>       | <span data-ttu-id="a0bab-152">Vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction.</span><span class="sxs-lookup"><span data-stu-id="a0bab-152">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="a0bab-153">Reportez-vous aux [options](#options) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="a0bab-153">See [options](#options) for details.</span></span>                                                          |
| `parameters`  | <span data-ttu-id="a0bab-154">tableau</span><span class="sxs-lookup"><span data-stu-id="a0bab-154">array</span></span>     | <span data-ttu-id="a0bab-155">Oui</span><span class="sxs-lookup"><span data-stu-id="a0bab-155">Yes</span></span>      | <span data-ttu-id="a0bab-156">Tableau qui définit les paramètres d’entrée de la fonction.</span><span class="sxs-lookup"><span data-stu-id="a0bab-156">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="a0bab-157">Pour plus [d’informations,](#parameters) voir les paramètres.</span><span class="sxs-lookup"><span data-stu-id="a0bab-157">See [parameters](#parameters) for details.</span></span>                                                                             |
| `result`      | <span data-ttu-id="a0bab-158">objet</span><span class="sxs-lookup"><span data-stu-id="a0bab-158">object</span></span>    | <span data-ttu-id="a0bab-159">Oui</span><span class="sxs-lookup"><span data-stu-id="a0bab-159">Yes</span></span>      | <span data-ttu-id="a0bab-160">Objet qui définit le type d’informations renvoyées par la fonction.</span><span class="sxs-lookup"><span data-stu-id="a0bab-160">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="a0bab-161">Reportez-vous au [résultat](#result) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="a0bab-161">See [result](#result) for details.</span></span>                                                                 |

### <a name="options"></a><span data-ttu-id="a0bab-162">options</span><span class="sxs-lookup"><span data-stu-id="a0bab-162">options</span></span>

<span data-ttu-id="a0bab-163">L’objet `options` vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction.</span><span class="sxs-lookup"><span data-stu-id="a0bab-163">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="a0bab-164">Le tableau suivant répertorie les propriétés de l’objet `options`.</span><span class="sxs-lookup"><span data-stu-id="a0bab-164">The following table lists the properties of the `options` object.</span></span>

| <span data-ttu-id="a0bab-165">Propriété</span><span class="sxs-lookup"><span data-stu-id="a0bab-165">Property</span></span>          | <span data-ttu-id="a0bab-166">Type de données</span><span class="sxs-lookup"><span data-stu-id="a0bab-166">Data type</span></span> | <span data-ttu-id="a0bab-167">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="a0bab-167">Required</span></span>                               | <span data-ttu-id="a0bab-168">Description</span><span class="sxs-lookup"><span data-stu-id="a0bab-168">Description</span></span> |
| :---------------- | :-------- | :------------------------------------- | :---------- |
| `cancelable`      | <span data-ttu-id="a0bab-169">boolean</span><span class="sxs-lookup"><span data-stu-id="a0bab-169">boolean</span></span>   | <span data-ttu-id="a0bab-170">Non</span><span class="sxs-lookup"><span data-stu-id="a0bab-170">No</span></span><br/><br/><span data-ttu-id="a0bab-171">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="a0bab-171">Default value is `false`.</span></span>  | <span data-ttu-id="a0bab-172">Si la valeur est `true`, Excel appelle le gestionnaire `CancelableInvocation` chaque fois que l’utilisateur effectue une action ayant pour effet d’annuler la fonction, par exemple, en déclenchant manuellement un recalcul ou en modifiant une cellule référencée par la fonction.</span><span class="sxs-lookup"><span data-stu-id="a0bab-172">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="a0bab-173">Les fonctions annulables sont généralement utilisées uniquement pour les fonctions asynchrones qui retournent un résultat unique et qui doivent gérer l’annulation d’une demande de données.</span><span class="sxs-lookup"><span data-stu-id="a0bab-173">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="a0bab-174">Une fonction ne peut pas utiliser les `stream` propriétés et les `cancelable` propriétés.</span><span class="sxs-lookup"><span data-stu-id="a0bab-174">A function can't use both the `stream` and `cancelable` properties.</span></span> |
| `requiresAddress` | <span data-ttu-id="a0bab-175">boolean</span><span class="sxs-lookup"><span data-stu-id="a0bab-175">boolean</span></span>   | <span data-ttu-id="a0bab-176">Non</span><span class="sxs-lookup"><span data-stu-id="a0bab-176">No</span></span> <br/><br/><span data-ttu-id="a0bab-177">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="a0bab-177">Default value is `false`.</span></span> | <span data-ttu-id="a0bab-178">Si `true` , votre fonction personnalisée peut accéder à l’adresse de la cellule qui l’a appelé.</span><span class="sxs-lookup"><span data-stu-id="a0bab-178">If `true`, your custom function can access the address of the cell that invoked it.</span></span> <span data-ttu-id="a0bab-179">La `address` propriété du paramètre [d’appel](custom-functions-parameter-options.md#invocation-parameter) contient l’adresse de la cellule qui a appelé votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="a0bab-179">The `address` property of the [invocation parameter](custom-functions-parameter-options.md#invocation-parameter) contains the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="a0bab-180">Une fonction ne peut pas utiliser les `stream` propriétés et les `requiresAddress` propriétés.</span><span class="sxs-lookup"><span data-stu-id="a0bab-180">A function can't use both the `stream` and `requiresAddress` properties.</span></span> |
| `requiresParameterAddresses` | <span data-ttu-id="a0bab-181">boolean</span><span class="sxs-lookup"><span data-stu-id="a0bab-181">boolean</span></span>   | <span data-ttu-id="a0bab-182">Non</span><span class="sxs-lookup"><span data-stu-id="a0bab-182">No</span></span> <br/><br/><span data-ttu-id="a0bab-183">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="a0bab-183">Default value is `false`.</span></span> | <span data-ttu-id="a0bab-184">Si `true` , votre fonction personnalisée peut accéder aux adresses des paramètres d’entrée de la fonction.</span><span class="sxs-lookup"><span data-stu-id="a0bab-184">If `true`, your custom function can access the addresses of the function's input parameters.</span></span> <span data-ttu-id="a0bab-185">Cette propriété doit être utilisée en association avec la propriété de l’objet de résultat et doit `dimensionality` être définie sur [](#result) `dimensionality` `matrix` .</span><span class="sxs-lookup"><span data-stu-id="a0bab-185">This property must be used in combination with the `dimensionality` property of the [result](#result) object, and `dimensionality` must be set to `matrix`.</span></span> <span data-ttu-id="a0bab-186">Pour [plus d’informations, voir Détecter l’adresse d’un](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) paramètre.</span><span class="sxs-lookup"><span data-stu-id="a0bab-186">See [Detect the address of a parameter](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) for more information.</span></span> |
| `stream`          | <span data-ttu-id="a0bab-187">boolean</span><span class="sxs-lookup"><span data-stu-id="a0bab-187">boolean</span></span>   | <span data-ttu-id="a0bab-188">Non</span><span class="sxs-lookup"><span data-stu-id="a0bab-188">No</span></span><br/><br/><span data-ttu-id="a0bab-189">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="a0bab-189">Default value is `false`.</span></span>  | <span data-ttu-id="a0bab-190">Si la valeur est `true`, la fonction peut envoyer une sortie à la cellule à plusieurs reprises, même en cas d’appel unique.</span><span class="sxs-lookup"><span data-stu-id="a0bab-190">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="a0bab-191">Cette option est utile pour des sources de données qui changent rapidement, telles que des valeurs boursières.</span><span class="sxs-lookup"><span data-stu-id="a0bab-191">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="a0bab-192">La fonction ne doit pas utiliser d’instruction `return`.</span><span class="sxs-lookup"><span data-stu-id="a0bab-192">The function should have no `return` statement.</span></span> <span data-ttu-id="a0bab-193">Au lieu de cela, la valeur obtenue est transmise en tant qu’argument de la méthode de rappel `StreamingInvocation.setResult`.</span><span class="sxs-lookup"><span data-stu-id="a0bab-193">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="a0bab-194">Pour plus d’informations, [voir Faire une fonction de diffusion en continu.](custom-functions-web-reqs.md#make-a-streaming-function)</span><span class="sxs-lookup"><span data-stu-id="a0bab-194">For more information, see [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
| `volatile`        | <span data-ttu-id="a0bab-195">boolean</span><span class="sxs-lookup"><span data-stu-id="a0bab-195">boolean</span></span>   | <span data-ttu-id="a0bab-196">Non</span><span class="sxs-lookup"><span data-stu-id="a0bab-196">No</span></span> <br/><br/><span data-ttu-id="a0bab-197">La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="a0bab-197">Default value is `false`.</span></span> | <span data-ttu-id="a0bab-198">Si , la fonction recalcule chaque fois que Excel recalcule, et non uniquement lorsque les valeurs dépendantes de la `true` formule ont changé.</span><span class="sxs-lookup"><span data-stu-id="a0bab-198">If `true`, the function recalculates each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="a0bab-199">Une fonction ne peut pas utiliser les `stream` propriétés et les `volatile` propriétés.</span><span class="sxs-lookup"><span data-stu-id="a0bab-199">A function can't use both the `stream` and `volatile` properties.</span></span> <span data-ttu-id="a0bab-200">Si les `stream` `volatile` propriétés et les propriétés sont définies sur , la propriété `true` volatile est ignorée.</span><span class="sxs-lookup"><span data-stu-id="a0bab-200">If the `stream` and `volatile` properties are both set to `true`, the volatile property will be ignored.</span></span> |

### <a name="parameters"></a><span data-ttu-id="a0bab-201">paramètres</span><span class="sxs-lookup"><span data-stu-id="a0bab-201">parameters</span></span>

<span data-ttu-id="a0bab-202">La propriété `parameters` est un tableau d’objets paramètre.</span><span class="sxs-lookup"><span data-stu-id="a0bab-202">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="a0bab-203">Le tableau suivant répertorie les propriétés de chaque objet.</span><span class="sxs-lookup"><span data-stu-id="a0bab-203">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="a0bab-204">Propriété</span><span class="sxs-lookup"><span data-stu-id="a0bab-204">Property</span></span>  |  <span data-ttu-id="a0bab-205">Type de données</span><span class="sxs-lookup"><span data-stu-id="a0bab-205">Data type</span></span>  |  <span data-ttu-id="a0bab-206">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="a0bab-206">Required</span></span>  |  <span data-ttu-id="a0bab-207">Description</span><span class="sxs-lookup"><span data-stu-id="a0bab-207">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="a0bab-208">string</span><span class="sxs-lookup"><span data-stu-id="a0bab-208">string</span></span>  |  <span data-ttu-id="a0bab-209">Non</span><span class="sxs-lookup"><span data-stu-id="a0bab-209">No</span></span> |  <span data-ttu-id="a0bab-210">Description du paramètre.</span><span class="sxs-lookup"><span data-stu-id="a0bab-210">A description of the parameter.</span></span> <span data-ttu-id="a0bab-211">Elle s’affiche dans Excel’IntelliSense.</span><span class="sxs-lookup"><span data-stu-id="a0bab-211">This is displayed in Excel's IntelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="a0bab-212">string</span><span class="sxs-lookup"><span data-stu-id="a0bab-212">string</span></span>  |  <span data-ttu-id="a0bab-213">Non</span><span class="sxs-lookup"><span data-stu-id="a0bab-213">No</span></span>  |  <span data-ttu-id="a0bab-214">Doit être `scalar` (une valeur autre qu’un tableau) ou (un tableau `matrix` à 2 dimensions).</span><span class="sxs-lookup"><span data-stu-id="a0bab-214">Must be either `scalar` (a non-array value) or `matrix` (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="a0bab-215">string</span><span class="sxs-lookup"><span data-stu-id="a0bab-215">string</span></span>  |  <span data-ttu-id="a0bab-216">Oui</span><span class="sxs-lookup"><span data-stu-id="a0bab-216">Yes</span></span>  |  <span data-ttu-id="a0bab-217">Le nom du paramètre.</span><span class="sxs-lookup"><span data-stu-id="a0bab-217">The name of the parameter.</span></span> <span data-ttu-id="a0bab-218">Ce nom s’affiche dans Excel’IntelliSense.</span><span class="sxs-lookup"><span data-stu-id="a0bab-218">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="a0bab-219">string</span><span class="sxs-lookup"><span data-stu-id="a0bab-219">string</span></span>  |  <span data-ttu-id="a0bab-220">Non</span><span class="sxs-lookup"><span data-stu-id="a0bab-220">No</span></span>  |  <span data-ttu-id="a0bab-221">Type de données du paramètre.</span><span class="sxs-lookup"><span data-stu-id="a0bab-221">The data type of the parameter.</span></span> <span data-ttu-id="a0bab-222">Peut être , ou , qui vous permet d’utiliser l’un des trois `boolean` `number` types `string` `any` précédents.</span><span class="sxs-lookup"><span data-stu-id="a0bab-222">Can be `boolean`, `number`, `string`, or `any`, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="a0bab-223">Si cette propriété n’est pas spécifiée, le type de données est par défaut `any` .</span><span class="sxs-lookup"><span data-stu-id="a0bab-223">If this property is not specified, the data type defaults to `any`.</span></span> |
|  `optional`  | <span data-ttu-id="a0bab-224">boolean</span><span class="sxs-lookup"><span data-stu-id="a0bab-224">boolean</span></span> | <span data-ttu-id="a0bab-225">Non</span><span class="sxs-lookup"><span data-stu-id="a0bab-225">No</span></span> | <span data-ttu-id="a0bab-226">Si la valeur est `true`, le paramètre est facultatif.</span><span class="sxs-lookup"><span data-stu-id="a0bab-226">If `true`, the parameter is optional.</span></span> |
|`repeating`| <span data-ttu-id="a0bab-227">boolean</span><span class="sxs-lookup"><span data-stu-id="a0bab-227">boolean</span></span> | <span data-ttu-id="a0bab-228">Non</span><span class="sxs-lookup"><span data-stu-id="a0bab-228">No</span></span> | <span data-ttu-id="a0bab-229">Si `true` , les paramètres sont remplis à partir d’un tableau spécifié.</span><span class="sxs-lookup"><span data-stu-id="a0bab-229">If `true`, parameters populate from a specified array.</span></span> <span data-ttu-id="a0bab-230">Notez que, par définition, tous les paramètres exexionnels sont considérés comme des paramètres facultatifs.</span><span class="sxs-lookup"><span data-stu-id="a0bab-230">Note that functions all repeating parameters are considered optional parameters by definition.</span></span>  |

### <a name="result"></a><span data-ttu-id="a0bab-231">résultat</span><span class="sxs-lookup"><span data-stu-id="a0bab-231">result</span></span>

<span data-ttu-id="a0bab-232">L’objet `result` définit le type des informations renvoyées par la fonction.</span><span class="sxs-lookup"><span data-stu-id="a0bab-232">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="a0bab-233">Le tableau suivant répertorie les propriétés de l’objet `result`.</span><span class="sxs-lookup"><span data-stu-id="a0bab-233">The following table lists the properties of the `result` object.</span></span>

| <span data-ttu-id="a0bab-234">Propriété</span><span class="sxs-lookup"><span data-stu-id="a0bab-234">Property</span></span>         | <span data-ttu-id="a0bab-235">Type de données</span><span class="sxs-lookup"><span data-stu-id="a0bab-235">Data type</span></span> | <span data-ttu-id="a0bab-236">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="a0bab-236">Required</span></span> | <span data-ttu-id="a0bab-237">Description</span><span class="sxs-lookup"><span data-stu-id="a0bab-237">Description</span></span>                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | <span data-ttu-id="a0bab-238">string</span><span class="sxs-lookup"><span data-stu-id="a0bab-238">string</span></span>    | <span data-ttu-id="a0bab-239">Non</span><span class="sxs-lookup"><span data-stu-id="a0bab-239">No</span></span>       | <span data-ttu-id="a0bab-240">Doit être `scalar` (une valeur autre qu’un tableau) ou (un tableau `matrix` à 2 dimensions).</span><span class="sxs-lookup"><span data-stu-id="a0bab-240">Must be either `scalar` (a non-array value) or `matrix` (a 2-dimensional array).</span></span> |
| `type` | <span data-ttu-id="a0bab-241">string</span><span class="sxs-lookup"><span data-stu-id="a0bab-241">string</span></span>    | <span data-ttu-id="a0bab-242">Non</span><span class="sxs-lookup"><span data-stu-id="a0bab-242">No</span></span>       | <span data-ttu-id="a0bab-243">Type de données du résultat.</span><span class="sxs-lookup"><span data-stu-id="a0bab-243">The data type of the result.</span></span> <span data-ttu-id="a0bab-244">Peut être , ou (ce qui vous permet d’utiliser l’un des trois `boolean` `number` types `string` `any` précédents).</span><span class="sxs-lookup"><span data-stu-id="a0bab-244">Can be `boolean`, `number`, `string`, or `any` (which allows you to use of any of the previous three types).</span></span> <span data-ttu-id="a0bab-245">Si cette propriété n’est pas spécifiée, le type de données est par défaut `any` .</span><span class="sxs-lookup"><span data-stu-id="a0bab-245">If this property is not specified, the data type defaults to `any`.</span></span> |

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="a0bab-246">Mappage des noms de fonction aux métadonnées JSON</span><span class="sxs-lookup"><span data-stu-id="a0bab-246">Associating function names with JSON metadata</span></span>

<span data-ttu-id="a0bab-247">Pour qu’une fonction fonctionne correctement, vous devez associer la propriété de la fonction `id` à l’implémentation JavaScript.</span><span class="sxs-lookup"><span data-stu-id="a0bab-247">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="a0bab-248">Assurez-vous qu’il existe une association, sinon la fonction ne sera pas enregistrée et ne peut pas être Excel.</span><span class="sxs-lookup"><span data-stu-id="a0bab-248">Make sure there is an association, otherwise the function won't be registered and isn't useable in Excel.</span></span> <span data-ttu-id="a0bab-249">L’exemple de code suivant montre comment faire en sorte que l’association utilise la `CustomFunctions.associate()` méthode.</span><span class="sxs-lookup"><span data-stu-id="a0bab-249">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="a0bab-250">L’exemple définit la fonction personnalisée `add` et associe à l’objet dans le fichier de métadonnées JSON où la valeur de la propriété`id`est **AJOUTER**.</span><span class="sxs-lookup"><span data-stu-id="a0bab-250">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

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

<span data-ttu-id="a0bab-251">Le code JSON suivant présente les métadonnées JSON associées au code JavaScript de la fonction personnalisée précédente.</span><span class="sxs-lookup"><span data-stu-id="a0bab-251">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

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

<span data-ttu-id="a0bab-252">N’oubliez pas les meilleures pratiques suivantes lors de la création de fonctions personnalisées dans votre fichier JavaScript et spécifiez les informations correspondantes dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="a0bab-252">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

- <span data-ttu-id="a0bab-253">Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété contient uniquement des points et des caractères alphanumériques.</span><span class="sxs-lookup"><span data-stu-id="a0bab-253">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

- <span data-ttu-id="a0bab-254">Dans le fichier de métadonnées JSON, vérifiez que la valeur de chaque `id` propriété est unique dans l’étendue du fichier.</span><span class="sxs-lookup"><span data-stu-id="a0bab-254">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="a0bab-255">Autrement dit, aucun objet fonction dans le fichier de métadonnées ne doit pas avoir la même`id`valeur.</span><span class="sxs-lookup"><span data-stu-id="a0bab-255">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

- <span data-ttu-id="a0bab-256">Ne modifiez pas la valeur d’une`id` propriété dans le fichier de métadonnées JSON après qu’elle ait été mappée à un nom de fonction JavaScript correspondante.</span><span class="sxs-lookup"><span data-stu-id="a0bab-256">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="a0bab-257">Vous pouvez modifier le nom de fonction que voient les utilisateurs finaux dans Excel en mettant à jour la `name` propriété dans le fichier de métadonnées JSON, mais vous ne devez jamais changer la valeur d’une `id` propriété après qu’elle a été établie.</span><span class="sxs-lookup"><span data-stu-id="a0bab-257">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

- <span data-ttu-id="a0bab-258">Dans le fichier JavaScript, spécifiez une association de fonction personnalisée à l’aide `CustomFunctions.associate` d’après chaque fonction.</span><span class="sxs-lookup"><span data-stu-id="a0bab-258">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="a0bab-259">L’exemple suivant montre les métadonnées JSON qui correspondent aux fonctions définies dans l’exemple de code JavaScript précédent.</span><span class="sxs-lookup"><span data-stu-id="a0bab-259">The following sample shows the JSON metadata that corresponds to the functions defined in the preceding JavaScript code sample.</span></span> <span data-ttu-id="a0bab-260">Les valeurs de propriété et les valeurs sont en minuscules, ce qui est une meilleure pratique lorsque vous `id` `name` décrivez vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="a0bab-260">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="a0bab-261">Vous devez ajouter ce JSON uniquement si vous préparez manuellement votre propre fichier JSON sans utiliser la génération automatique.</span><span class="sxs-lookup"><span data-stu-id="a0bab-261">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="a0bab-262">Pour plus d’informations sur la génération automatique, voir [autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="a0bab-262">For more information on autogeneration, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="a0bab-263">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="a0bab-263">Next steps</span></span>

<span data-ttu-id="a0bab-264">Découvrez les [meilleures pratiques pour nommer](custom-functions-naming.md) votre [](custom-functions-localize.md) fonction ou découvrir comment la localiser à l’aide de la méthode JSON manuscrite précédemment décrite.</span><span class="sxs-lookup"><span data-stu-id="a0bab-264">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="a0bab-265">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a0bab-265">See also</span></span>

- [<span data-ttu-id="a0bab-266">Générer automatiquement des métadonnées JSON pour des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="a0bab-266">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="a0bab-267">Options des paramètres de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="a0bab-267">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
- [<span data-ttu-id="a0bab-268">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="a0bab-268">Create custom functions in Excel</span></span>](custom-functions-overview.md)
