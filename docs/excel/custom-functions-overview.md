---
ms.date: 09/27/2018
description: Créez une fonction personnalisée dans Excel à l’aide de JavaScript.
title: Créer des fonctions personnalisées dans Excel (Aperçu)
ms.openlocfilehash: 98e418f843f6f5574088cea9c7393afc4a42060b
ms.sourcegitcommit: 1852ae367de53deb91d03ca55d16eb69709340d3
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/29/2018
ms.locfileid: "25348800"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="03d09-103">Créer des fonctions personnalisées dans Excel (aperçu)</span><span class="sxs-lookup"><span data-stu-id="03d09-103">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="03d09-p101">Les fonctions personnalisées permettent aux développeurs d'ajouter de nouvelles fonctions à Excel en définissant ces fonctions dans JavaScript comme partie d’un complément. Les utilisateurs d'Excel peuvent accéder à des fonctions personnalisées comme n'importe quelle fonction native d'Excel, telle que `SUM()`. Cet article explique comment créer des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="03d09-p101">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions like any other native function in Excel (such as `SUM()`). This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="03d09-107">L’illustration suivante montre un utilisateur final insérant une fonction personnalisée dans une cellule d’une feuille de calcul Excel.</span><span class="sxs-lookup"><span data-stu-id="03d09-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="03d09-108">La fonction personnalisée `CONTOSO.ADD42` est conçue pour ajouter 42 à la paire de nombres spécifiée par l’utilisateur comme paramètres d’entrée de la fonction.</span><span class="sxs-lookup"><span data-stu-id="03d09-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="03d09-109">Le code suivant définit la fonction personnalisée `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="03d09-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="03d09-110">Plus loin dans cet article, la section [Problèmes connus](#known-issues) indique les limites actuelles des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="03d09-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="03d09-111">Composants d’un projet de complément de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="03d09-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="03d09-112">Si vous utilisez le [générateur de Yo Office](https://github.com/OfficeDev/generator-office) pour créer un projet de complément de fonctions personnalisées Excel, vous verrez les fichiers suivants dans le projet que le générateur crée :</span><span class="sxs-lookup"><span data-stu-id="03d09-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="03d09-113">Fichier</span><span class="sxs-lookup"><span data-stu-id="03d09-113">File</span></span> | <span data-ttu-id="03d09-114">Format de fichier</span><span class="sxs-lookup"><span data-stu-id="03d09-114">File format</span></span> | <span data-ttu-id="03d09-115">Description</span><span class="sxs-lookup"><span data-stu-id="03d09-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="03d09-116">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="03d09-116">**./src/customfunctions.js**</span></span><br/><span data-ttu-id="03d09-117">ou</span><span class="sxs-lookup"><span data-stu-id="03d09-117">or</span></span><br/><span data-ttu-id="03d09-118">**./src/customfunctions.ts**</span><span class="sxs-lookup"><span data-stu-id="03d09-118">**./src/customfunctions.ts**</span></span> | <span data-ttu-id="03d09-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="03d09-119">JavaScript</span></span><br/><span data-ttu-id="03d09-120">ou</span><span class="sxs-lookup"><span data-stu-id="03d09-120">or</span></span><br/><span data-ttu-id="03d09-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="03d09-121">TypeScript</span></span> | <span data-ttu-id="03d09-122">Contient le code qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="03d09-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="03d09-123">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="03d09-123">**./config/customfunctions.json**</span></span> | <span data-ttu-id="03d09-124">JSON</span><span class="sxs-lookup"><span data-stu-id="03d09-124">JSON</span></span> | <span data-ttu-id="03d09-125">Contient des métadonnées qui décrivent les fonctions personnalisées et permettent à Excel d'enregistrer les fonctions personnalisées et de les mettre à la disposition des utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="03d09-125">Contains metadata that describes custom functions and enables Excel to register the custom functions in order to make them available to end-users.</span></span> |
| <span data-ttu-id="03d09-126">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="03d09-126">**./index.html**</span></span> | <span data-ttu-id="03d09-127">HTML</span><span class="sxs-lookup"><span data-stu-id="03d09-127">HTML</span></span> | <span data-ttu-id="03d09-128">Fournit une référence de &lt;script&gt; pour le fichier JavaScript qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="03d09-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="03d09-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="03d09-129">**Manifest.xml**</span></span> | <span data-ttu-id="03d09-130">XML</span><span class="sxs-lookup"><span data-stu-id="03d09-130">XML</span></span> | <span data-ttu-id="03d09-131">Spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers JavaScript, JSON et HTML répertoriés précédemment dans ce tableau.</span><span class="sxs-lookup"><span data-stu-id="03d09-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="03d09-132">Les sections suivantes fournissent plus d’informations sur ces fichiers.</span><span class="sxs-lookup"><span data-stu-id="03d09-132">The following sections provide more information about those changes.</span></span>

### <a name="script-file"></a><span data-ttu-id="03d09-133">Fichier de script</span><span class="sxs-lookup"><span data-stu-id="03d09-133">Script file</span></span> 

<span data-ttu-id="03d09-134">Le fichier de script (**./src/customfunctions.js** ou **./src/customfunctions.ts** dans le projet que le générateur de Yo Office crée) contient le code qui définit les fonctions personnalisées et mappe les noms des fonctions personnalisées aux objets du [fichier de métadonnées JSON](#json-metadata-file).</span><span class="sxs-lookup"><span data-stu-id="03d09-134">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="03d09-135">Par exemple, le code suivant définit les fonctions personnalisées `add` et `increment`, puis spécifie les informations de mappage pour les deux fonctions.</span><span class="sxs-lookup"><span data-stu-id="03d09-135">For example, the following code defines the custom functions `add` and `increment` and then specifies mapping information for both functions.</span></span> <span data-ttu-id="03d09-136">La fonction `add` est mappée à l'objet dans le fichier de métadonnées JSON où la valeur de la propriété `id` est **ADD**, et la fonction `increment` est mappée à l'objet dans le fichier de métadonnées où la valeur de la propriété `id` est **INCREMENT**.</span><span class="sxs-lookup"><span data-stu-id="03d09-136">The `add` function is mapped to the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is mapped to the object in the metadata file where the value of the `id` property is **INCREMENT**.</span></span> <span data-ttu-id="03d09-137">Pour plus d’informations sur le mappage des noms de fonction dans le fichier de script aux objets dans le fichier de métadonnées JSON, reportez-vous à la rubrique [Meilleures pratiques des fonctions personnalisées](custom-functions-best-practices.md#mapping-function-names-to-json-metadata).</span><span class="sxs-lookup"><span data-stu-id="03d09-137">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about mapping function names in the script file to objects in the JSON metadata file.</span></span>

```js
function add(first, second){
  return first + second;
}

function increment(incrementBy, callback) {
  var result = 0;
  var timer = setInterval(function() {
    result += incrementBy;
    callback.setResult(result);
  }, 1000);

  callback.onCanceled = function() {
    clearInterval(timer);
  };
}

// map `id` values in the JSON metadata file to the JavaScript function names
CustomFunctionMappings.ADD = add;
CustomFunctionMappings.INCREMENT = increment;
```

### <a name="json-metadata-file"></a><span data-ttu-id="03d09-138">Fichier de métadonnées JSON</span><span class="sxs-lookup"><span data-stu-id="03d09-138">JSON metadata file</span></span> 

<span data-ttu-id="03d09-139">Le fichier de métadonnées des fonctions personnalisées (**./config/customfunctions.json** dans le projet que le générateur de Yo Office crée) fournit les informations dont Excel a besoin pour enregistrer les fonctions personnalisées et les rendre disponibles aux utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="03d09-139">The custom functions metadata file (**./config/customfunctions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users.</span></span> <span data-ttu-id="03d09-140">Les fonctions personnalisées sont enregistrées lorsqu’un utilisateur exécute un complément pour la première fois.</span><span class="sxs-lookup"><span data-stu-id="03d09-140">The custom functions are registered when a user runs the add-in for the first time.</span></span> <span data-ttu-id="03d09-141">Après cela, elles sont disponibles pour cet utilisateur dans tous les classeurs (autrement dit, pas seulement dans le classeur dans lequel le complément a été exécuté pour la première fois.)</span><span class="sxs-lookup"><span data-stu-id="03d09-141">After that, they are available, for that same user, in all workbooks (not only the one where the add-in ran initially.)</span></span>

> [!TIP]
> <span data-ttu-id="03d09-142">Parmi les paramètres de serveur sur le serveur qui héberge le fichier JSON, [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) doit être activé pour que les fonctions personnalisées fonctionnent correctement dans Excel Online.</span><span class="sxs-lookup"><span data-stu-id="03d09-142">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="03d09-143">Le code suivant dans **customfunctions.json** spécifie les métadonnées de la fonction `add` et de la fonction `increment` décrites précédemment.</span><span class="sxs-lookup"><span data-stu-id="03d09-143">The following code in **customfunctions.json** specifies the metadata for the `add` function that was described previously in this article.</span></span> <span data-ttu-id="03d09-144">Le tableau qui suit cet échantillon de code fournit des informations détaillées sur les propriétés individuelles dans cet objet JSON.</span><span class="sxs-lookup"><span data-stu-id="03d09-144">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span> <span data-ttu-id="03d09-145">Reportez-vous à la rubrique [Meilleures pratiques des fonctions personnalisées](custom-functions-best-practices.md#mapping-function-names-to-json-metadata), pour plus d'informations sur la spécification de la valeur des propriétés `id` et `name` dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="03d09-145">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      "description": "Add two numbers",
      "helpUrl": "http://www.contoso.com",
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
      "id": "INCREMENT",
      "name": "INCREMENT",
      "description": "Periodically increment a value",
      "helpUrl": "http://www.contoso.com",
      "result": {
          "type": "number",
          "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "increment",
            "description": "Amount to increment",
            "type": "number",
            "dimensionality": "scalar"
        }
    ],
    "options": {
        "cancelable": true,
        "stream": true
      }
    }
  ]
}
```

<span data-ttu-id="03d09-146">Le tableau suivant répertorie les propriétés qui sont généralement présentes dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="03d09-146">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="03d09-147">Pour plus d'informations détaillées sur le fichier de métadonnées JSON, voir [Métadonnées des fonctions personnalisées](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="03d09-147">For more detailed information about the JSON metadata file, including options not used in the previous example, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="03d09-148">Propriété</span><span class="sxs-lookup"><span data-stu-id="03d09-148">Property</span></span>  | <span data-ttu-id="03d09-149">Description</span><span class="sxs-lookup"><span data-stu-id="03d09-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="03d09-150">ID unique de la fonction.</span><span class="sxs-lookup"><span data-stu-id="03d09-150">A unique ID for the group.</span></span> <span data-ttu-id="03d09-151">Cet ID ne doit pas être modifié après sa définition.</span><span class="sxs-lookup"><span data-stu-id="03d09-151">This ID should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="03d09-152">Nom de la fonction que l’utilisateur final voit dans Excel.</span><span class="sxs-lookup"><span data-stu-id="03d09-152">Name of the function that the end user sees in Excel.</span></span> <span data-ttu-id="03d09-153">Dans Excel, ce nom de fonction sera préfixé par l'espace de noms des fonctions personnalisées, spécifié dans le [fichier manifeste XML](#manifest-file).</span><span class="sxs-lookup"><span data-stu-id="03d09-153">In the autocomplete menu, this value will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `helpUrl` | <span data-ttu-id="03d09-154">URL de la page qui s’affiche lorsqu’un utilisateur demande de l’aide.</span><span class="sxs-lookup"><span data-stu-id="03d09-154">Url for a page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="03d09-155">Décrit ce que fait la fonction.</span><span class="sxs-lookup"><span data-stu-id="03d09-155">Describes what the function does.</span></span> <span data-ttu-id="03d09-156">Cette valeur s’affiche comme une info-bulle lorsque la fonction est l’élément sélectionné dans le menu de saisie semi-automatique dans Excel.</span><span class="sxs-lookup"><span data-stu-id="03d09-156">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="03d09-157">Objet qui définit le type de l’information renvoyée par la fonction.</span><span class="sxs-lookup"><span data-stu-id="03d09-157">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="03d09-158">La valeur de la propriété enfant `type` peut être **string**, **number**ou **boolean**.</span><span class="sxs-lookup"><span data-stu-id="03d09-158">The value of the `type` child property can be **string**, **number**, or **boolean**.</span></span> <span data-ttu-id="03d09-159">La valeur de la propriété enfant `dimensionality` peut être **scalaire** ou **matrice** (un tableau à deux dimensions des valeurs de `type` spécifié).</span><span class="sxs-lookup"><span data-stu-id="03d09-159">The `dimensionality` property can be \*\*\*\* or \*\*\*\* (a two-dimensional array of values of the specified `type`.)</span></span> |
| `parameters` | <span data-ttu-id="03d09-160">Tableau qui définit les paramètres d’entrée de la fonction.</span><span class="sxs-lookup"><span data-stu-id="03d09-160">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="03d09-161">Les propriétés enfants `name` et `description` apparaissent dans l’intelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="03d09-161">The `name` and `description` child properties are used in the Excel intellisense.</span></span> <span data-ttu-id="03d09-162">La valeur de la propriété enfant`type` peut être une **chaîne**, un **nombre**, ou une valeur **booléenne**.</span><span class="sxs-lookup"><span data-stu-id="03d09-162">The value of the `type` child property can be **string**, **number**, or **boolean**.</span></span> <span data-ttu-id="03d09-163">La valeur de la propriété enfant `dimensionality` peut être **scalaire** ou **matrice** (un tableau à deux dimensions des valeurs de `type` spécifié).</span><span class="sxs-lookup"><span data-stu-id="03d09-163">The `dimensionality` property can be \*\*\*\* or \*\*\*\* (a two-dimensional array of values of the specified `type`.)</span></span> |
| `options` | <span data-ttu-id="03d09-164">Vous permet de personnaliser certains aspects de la façon dont Excel exécute la fonction et quand.</span><span class="sxs-lookup"><span data-stu-id="03d09-164">The  property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="03d09-165">Pour plus d’informations sur l’utilisation de cette propriété, voir [Fonctions de flux](#streamed-functions) et [Annulation d'une fonction](#canceling-a-function), plus loin dans cet article.</span><span class="sxs-lookup"><span data-stu-id="03d09-165">For more information about how this property can be used, see [Streamed functions](#streamed-functions) and [Cancellation](#canceling-a-function) later in this article.</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="03d09-166">Fichier manifeste</span><span class="sxs-lookup"><span data-stu-id="03d09-166">Manifest file</span></span>

<span data-ttu-id="03d09-167">Le fichier manifeste XML pour un complément qui définit les fonctions personnalisées (**./manifest.xml** dans le projet que le générateur de Yo Office crée) spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers JavaScript, JSON et HTML.</span><span class="sxs-lookup"><span data-stu-id="03d09-167">The XML manifest file for an add-in that defines custom functions specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="03d09-168">La balise XML suivante montre un exemple des éléments `<ExtensionPoint>` et `<Resources>` que vous devez inclure dans le manifeste d'un complément pour activer les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="03d09-168">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest in order to enable Excel to run custom functions.</span></span>  

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="JS-URL" /> <!--resid points to location of JavaScript file-->
                    </Script>
                    <Page>
                        <SourceLocation resid="HTML-URL"/> <!--resid points to location of HTML file-->
                    </Page>
                    <Metadata>
                        <SourceLocation resid="JSON-URL" /> <!--resid points to location of JSON file-->
                    </Metadata>
                    <Namespace resid="namespace" />
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>
    <Resources>
        <bt:Urls>
            <bt:Url id="JSON-URL" DefaultValue="http://127.0.0.1:8080/customfunctions.json" /> <!--specifies the location of your JSON file-->
            <bt:Url id="JS-URL" DefaultValue="http://127.0.0.1:8080/customfunctions.js" /> <!--specifies the location of your JavaScript file-->
            <bt:Url id="HTML-URL" DefaultValue="http://127.0.0.1:8080/index.html" /> <!--specifies the location of your HTML file-->
        </bt:Urls>
        <bt:ShortStrings>
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. -->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="03d09-169">Les fonctions dans Excel sont précédées par l’espace de noms spécifié dans votre fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="03d09-169">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="03d09-170">L’espace de noms d’une fonction précède le nom de la fonction, et ils sont séparés par un point.</span><span class="sxs-lookup"><span data-stu-id="03d09-170">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="03d09-171">Par exemple, pour appeler la fonction `ADD42` dans la cellule d’une feuille de calcul Excel, vous devez taper `=CONTOSO.ADD42`, puisque CONTOSO est l’espace de noms et `ADD42` est le nom de la fonction spécifiée dans le fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="03d09-171">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because CONTOSO is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="03d09-172">L’espace de noms est destiné à être utilisé comme identificateur pour votre entreprise ou le complément.</span><span class="sxs-lookup"><span data-stu-id="03d09-172">The prefix is intended to be used as an identifier for your add-in.</span></span> 

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="03d09-173">Fonctions qui renvoient des données provenant de sources externes</span><span class="sxs-lookup"><span data-stu-id="03d09-173">Functions that return data from external sources</span></span>

<span data-ttu-id="03d09-174">Si une fonction personnalisée récupère les données d’une source externe comme le Web, elle doit :</span><span class="sxs-lookup"><span data-stu-id="03d09-174">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="03d09-175">Renvoyer une promesse JavaScript à Excel.</span><span class="sxs-lookup"><span data-stu-id="03d09-175">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="03d09-176">Résoudre la promesse avec la valeur finale en utilisant la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="03d09-176">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="03d09-177">Les fonctions personnalisées affichent un résultat temporaire `#GETTING_DATA` dans la cellule pendant qu’Excel attend le résultat final.</span><span class="sxs-lookup"><span data-stu-id="03d09-177">Asynchronous functions display a `#GETTING_DATA` temporary error in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="03d09-178">Les utilisateurs peuvent interagir normalement avec le reste de la feuille de calcul tout en attendant le résultat.</span><span class="sxs-lookup"><span data-stu-id="03d09-178">Users can interact normally with the rest of the spreadsheet while they wait for the result.</span></span>

<span data-ttu-id="03d09-179">Dans l’échantillon de code suivant, la fonction personnalisée `getTemperature()` récupère la température actuelle d’un thermomètre.</span><span class="sxs-lookup"><span data-stu-id="03d09-179">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer.</span></span> <span data-ttu-id="03d09-180">Remarquez que `sendWebRequest` est une fonction hypothétique (non spécifiée ici) qui utilise [XHR](custom-functions-runtime.md#xhr) pour appeler un service web de température.</span><span class="sxs-lookup"><span data-stu-id="03d09-180">Note that `sendWebRequest` is a hypothetical function, not specified here, that uses XHR to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a><span data-ttu-id="03d09-181">Fonctions de flux</span><span class="sxs-lookup"><span data-stu-id="03d09-181">Streamed functions</span></span>

<span data-ttu-id="03d09-182">Les fonctions de flux personnalisées vous permettent de transmettre des données aux cellules de manière répétée au fil du temps, sans qu'un utilisateur ait à demander explicitement une actualisation des données.</span><span class="sxs-lookup"><span data-stu-id="03d09-182">Streamed custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request recalculation.</span></span> <span data-ttu-id="03d09-183">L’échantillon de code suivant est une fonction personnalisée qui ajoute un nombre au résultat toutes les secondes.</span><span class="sxs-lookup"><span data-stu-id="03d09-183">The following example is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="03d09-184">Tenez compte des informations suivantes relatives à ce code :</span><span class="sxs-lookup"><span data-stu-id="03d09-184">Note the following about this code:</span></span>

- <span data-ttu-id="03d09-185">Excel affiche automatiquement chaque nouvelle valeur en utilisant le rappel `setResult`.</span><span class="sxs-lookup"><span data-stu-id="03d09-185">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="03d09-186">Le second paramètre d’entrée, `handler`, n’est pas affiché pour les utilisateurs finaux dans Excel lorsqu’ils sélectionnent la fonction à partir du menu de saisie semi-automatique.</span><span class="sxs-lookup"><span data-stu-id="03d09-186">The second input parameter, `handler`, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>

- <span data-ttu-id="03d09-187">Le rappel `onCanceled` définit la fonction qui s’exécute lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="03d09-187">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="03d09-188">Vous devez implémenter un gestionnaire d'annulation comme celui-ci pour toute fonction de flux.</span><span class="sxs-lookup"><span data-stu-id="03d09-188">You must implement a cancellation handler like this for any streamed function.</span></span> <span data-ttu-id="03d09-189">Pour plus d’informations, voir [Annulation d’une fonction](#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="03d09-189">For more information, see [Canceling a function](#canceling-a-function).</span></span> 

```js
function incrementValue(increment, handler){
  var result = 0;
  setInterval(function(){
    result += increment;
    handler.setResult(result);
  }, 1000);

  handler.onCanceled = function(){
    clearInterval(timer);
  }
}
```

<span data-ttu-id="03d09-190">Lorsque vous spécifiez des métadonnées pour une fonction de flux dans le fichier de métadonnées JSON, vous devez définir les propriétés `"cancelable": true` et `"stream": true` dans l'objet `options`, comme illustré dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="03d09-190">When you specify metadata for a streamed function in the JSON metadata file, you must set the properties `"cancelable": true` and `"stream": true` within the `options` object, as shown in the following example.</span></span>

```json
{
  "id": "INCREMENT",
  "name": "INCREMENT",
  "description": "Periodically increment a value",
  "helpUrl": "http://www.contoso.com",
  "result": {
    "type": "number",
    "dimensionality": "scalar"
  },
  "parameters": [
    {
      "name": "increment",
      "description": "Amount to increment",
      "type": "number",
      "dimensionality": "scalar"
    }
  ],
  "options": {
    "cancelable": true,
    "stream": true
  }
}
```

## <a name="canceling-a-function"></a><span data-ttu-id="03d09-191">Annulation d’une fonction</span><span class="sxs-lookup"><span data-stu-id="03d09-191">Canceling a function</span></span>

<span data-ttu-id="03d09-192">Dans certains cas, vous devrez peut-être annuler l’exécution d’une fonction personnalisée en flux continu pour réduire la consommation de la bande passante, de la mémoire et de la charge processeur.</span><span class="sxs-lookup"><span data-stu-id="03d09-192">In some situations, you may need to cancel the execution of a streamed custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="03d09-193">Excel annule l’exécution d’une fonction dans les situations suivantes :</span><span class="sxs-lookup"><span data-stu-id="03d09-193">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="03d09-194">Quand l’utilisateur modifie ou supprime une cellule qui fait référence à la fonction.</span><span class="sxs-lookup"><span data-stu-id="03d09-194">The user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="03d09-195">Quand un des arguments (entrées) de la fonction est modifié.</span><span class="sxs-lookup"><span data-stu-id="03d09-195">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="03d09-196">Dans ce cas, un nouvel appel de fonction est déclenché suite à l'annulation.</span><span class="sxs-lookup"><span data-stu-id="03d09-196">In this case, a new function call is triggered in addition to the cancelation.</span></span>

- <span data-ttu-id="03d09-197">Lorsque l’utilisateur déclenche manuellement un nouveau calcul.</span><span class="sxs-lookup"><span data-stu-id="03d09-197">When the user triggers recalculation manually.</span></span> <span data-ttu-id="03d09-198">Dans ce cas, un nouvel appel de fonction est déclenché suite à l'annulation.</span><span class="sxs-lookup"><span data-stu-id="03d09-198">In this case, a new function call is triggered in addition to the cancelation.</span></span>

<span data-ttu-id="03d09-199">Pour activer la possibilité d’annuler une fonction, vous devez implémenter un gestionnaire d’annulation dans la fonction JavaScript et spécifier la propriété `"cancelable": true`   dans l'objet `options`   dans les métadonnées JSON qui décrit la fonction.</span><span class="sxs-lookup"><span data-stu-id="03d09-199">To enable the ability to cancel a function, you must implement a cancellation handler within the JavaScript function and specify the property `"cancelable": true` within the `options` object in the JSON metadata that describes the function.</span></span> <span data-ttu-id="03d09-200">Les échantillons de code dans la section précédente de cet article fournissent un exemple de ces techniques.</span><span class="sxs-lookup"><span data-stu-id="03d09-200">The code samples in the previous section of this article provide an example of these techniques.</span></span>

## <a name="saving-and-sharing-state"></a><span data-ttu-id="03d09-201">Enregistrement et partage de l'état</span><span class="sxs-lookup"><span data-stu-id="03d09-201">Saving and sharing state</span></span>

<span data-ttu-id="03d09-202">Les fonctions personnalisées peuvent enregistrer des données dans des variables JavaScript globales.</span><span class="sxs-lookup"><span data-stu-id="03d09-202">Custom functions can save data in global JavaScript variables.</span></span> <span data-ttu-id="03d09-203">Lors d’appels ultérieurs, votre fonction personnalisée pourra utiliser les valeurs enregistrées dans ces variables.</span><span class="sxs-lookup"><span data-stu-id="03d09-203">In subsequent calls, your custom function may use the values saved in these variables.</span></span> <span data-ttu-id="03d09-204">L'état enregistré est utile lorsque les utilisateurs ajoutent la même fonction personnalisée à plusieurs cellules, car toutes les instances de la fonction peuvent partager l'état.</span><span class="sxs-lookup"><span data-stu-id="03d09-204">Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state.</span></span> <span data-ttu-id="03d09-205">Par exemple, vous pouvez enregistrer les données renvoyées par un appel à une ressource web pour éviter de passer des appels supplémentaires à la même ressource web.</span><span class="sxs-lookup"><span data-stu-id="03d09-205">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="03d09-206">L'échantillon de code suivant montre l'implémentation d'une fonction de flux de température qui enregistre l'état de manière globale.</span><span class="sxs-lookup"><span data-stu-id="03d09-206">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="03d09-207">Tenez compte des informations suivantes relatives à ce code :</span><span class="sxs-lookup"><span data-stu-id="03d09-207">Note the following about this code:</span></span>

- <span data-ttu-id="03d09-208">`refreshTemperature` ,est une fonction de flux qui chaque seconde, lit la température d’un thermomètre spécifique.</span><span class="sxs-lookup"><span data-stu-id="03d09-208">`refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second.</span></span> <span data-ttu-id="03d09-209">Les nouvelles températures sont enregistrées dans la variable `savedTemperatures`, mais ne mettent pas directement à jour la valeur de la cellule.</span><span class="sxs-lookup"><span data-stu-id="03d09-209">New temperatures are saved in the `savedTemperatures` variable, but does not directly update the cell value.</span></span> <span data-ttu-id="03d09-210">Elles ne doivent pas être appelées directement à partir d'une cellule de feuille de calcul, *de sorte qu'elles ne sont pas enregistrées dans le fichier JSON*.</span><span class="sxs-lookup"><span data-stu-id="03d09-210">It should not be directly called from a worksheet cell, *so it is not registered in the JSON file*.</span></span>

- <span data-ttu-id="03d09-211">`streamTemperature` met à jour les valeurs de température affichées dans la cellule, chaque seconde, et utilise une variable `savedTemperatures` comme source de données.</span><span class="sxs-lookup"><span data-stu-id="03d09-211">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span> <span data-ttu-id="03d09-212">Elles doivent être enregistrées dans le fichier JSON et nommées en lettres majuscules, `STREAMTEMPERATURE`.</span><span class="sxs-lookup"><span data-stu-id="03d09-212">It must be registered in the JSON file, and named with all upper-case letters, `STREAMTEMPERATURE`.</span></span>

- <span data-ttu-id="03d09-213">Les utilisateurs peuvent appeler `streamTemperature` à partir de plusieurs cellules dans l’interface utilisateur Excel.</span><span class="sxs-lookup"><span data-stu-id="03d09-213">Users may call `streamTemperature` from several cells in the Excel UI.</span></span> <span data-ttu-id="03d09-214">Chaque appel lit des données depuis la même variable `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="03d09-214">Each call reads data from the same `savedTemperatures` variable.</span></span>

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
  if(!savedTemperatures[thermometerID]){
    refreshTemperatures(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
  }

  function getNextTemperature(){
    handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
    setTimeout(getNextTemperature, 1000); // Wait 1 second before updating Excel again.
  }
  getNextTemperature();
}

function refreshTemperature(thermometerID){
  sendWebRequest(thermometerID, function(data){
    savedTemperatures[thermometerID] = data.temperature;
  });
  setTimeout(function(){
    refreshTemperature(thermometerID);
  }, 1000); // Wait 1 second before reading the thermometer again, and then update the saved temperature of thermometerID.
}
```

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="03d09-215">Utilisation des plages de données</span><span class="sxs-lookup"><span data-stu-id="03d09-215">Working with ranges of data</span></span>

<span data-ttu-id="03d09-216">Votre fonction personnalisée peut accepter une plage de données comme paramètre d’entrée, ou elle peut renvoyer une plage de données.</span><span class="sxs-lookup"><span data-stu-id="03d09-216">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="03d09-217">En JavaScript, une plage de données est représentée sous la forme d’un tableau à deux dimensions.</span><span class="sxs-lookup"><span data-stu-id="03d09-217">In JavaScript, a range of data is represented as a 2-dimensional array.</span></span>

<span data-ttu-id="03d09-218">Par exemple, supposons que votre fonction renvoie la deuxième valeur la plus élevée prise dans une plage de nombres stockés dans Excel.</span><span class="sxs-lookup"><span data-stu-id="03d09-218">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="03d09-219">La fonction suivante accepte le paramètre `values`, qui est de type `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="03d09-219">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="03d09-220">Notez que dans les métadonnées JSON de cette fonction, vous devez définir la propriété `type` du paramètre à `matrix`.</span><span class="sxs-lookup"><span data-stu-id="03d09-220">Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.</span></span>

```js
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
  for(var i = 0; i < values.length; i++){
    for(var j = 1; j < values[i].length; j++){
      if(values[i][j] >= highest){
        secondHighest = highest;
        highest = values[i][j];
      }
      else if(values[i][j] >= secondHighest){
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## <a name="handling-errors"></a><span data-ttu-id="03d09-221">Gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="03d09-221">Handling errors</span></span>

<span data-ttu-id="03d09-222">Lorsque vous créez un complément qui définit des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution.</span><span class="sxs-lookup"><span data-stu-id="03d09-222">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="03d09-223">La gestion des erreurs pour les fonctions personnalisées est identique à la [gestion des erreurs pour l’API JavaScript d’Excel dans son ensemble](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="03d09-223">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="03d09-224">Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.</span><span class="sxs-lookup"><span data-stu-id="03d09-224">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(x) {
  let url = "https://www.contoso.com/comments/" + x;

  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then((json) => {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## <a name="known-issues"></a><span data-ttu-id="03d09-225">Problèmes connus</span><span class="sxs-lookup"><span data-stu-id="03d09-225">Known issues</span></span>

- <span data-ttu-id="03d09-226">Les descriptions de paramètre et les URL d’aide ne sont pas encore utilisées par Excel.</span><span class="sxs-lookup"><span data-stu-id="03d09-226">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="03d09-227">Les fonctions personnalisées ne sont actuellement pas disponibles sur Excel pour les clients mobiles.</span><span class="sxs-lookup"><span data-stu-id="03d09-227">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="03d09-228">Les fonctions volatiles (celles qui recalculent automatiquement lorsque des modifications de données indépendantes sont effectuées dans la feuille de calcul) ne sont pas encore prises en charge.</span><span class="sxs-lookup"><span data-stu-id="03d09-228">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="03d09-229">Le déploiement via le portail d'administration Office 365 et AppSource n'est pas encore activé.</span><span class="sxs-lookup"><span data-stu-id="03d09-229">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="03d09-230">Les fonctions personnalisées dans Excel Online peuvent cesser de fonctionner pendant une session après une période d'inactivité.</span><span class="sxs-lookup"><span data-stu-id="03d09-230">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="03d09-231">Actualisez la page du navigateur (F5) et entrez à nouveau une fonction personnalisée pour restaurer la fonction.</span><span class="sxs-lookup"><span data-stu-id="03d09-231">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="03d09-232">Il est possible d’avoir le résultat temporaire **#GETTING_DATA** dans la ou les cellules d’une feuille de calcul si vous avez plusieurs compléments s’exécutant dans Microsoft Excel pour Windows.</span><span class="sxs-lookup"><span data-stu-id="03d09-232">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows.</span></span> <span data-ttu-id="03d09-233">Fermez toutes les fenêtres Excel et redémarrez Excel.</span><span class="sxs-lookup"><span data-stu-id="03d09-233">Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="03d09-234">Des outils de débogage spécifiques pour les fonctions personnalisées pourraient devenir disponibles à l’avenir.</span><span class="sxs-lookup"><span data-stu-id="03d09-234">Debugging tools specifically for custom functions may be available in the future.</span></span> <span data-ttu-id="03d09-235">En attendant, vous pouvez déboguer sur Excel Online à l’aide des outils de développement F12.</span><span class="sxs-lookup"><span data-stu-id="03d09-235">In the meantime, you can debug on Excel Online using F12 developer tools.</span></span> <span data-ttu-id="03d09-236">Voir plus de détails dans [Meilleures pratiques pour les fonctions personnalisées](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="03d09-236">See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>

## <a name="changelog"></a><span data-ttu-id="03d09-237">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="03d09-237">Changelog</span></span>

- <span data-ttu-id="03d09-238">**7 novembre 2017 :** mise à disposition\* de la préversion des fonctions personnalisées et d'exemples</span><span class="sxs-lookup"><span data-stu-id="03d09-238">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="03d09-239">**20 novembre 2017**: correction du bogue de compatibilité pour les utilisateurs de la version 8801 et ultérieure</span><span class="sxs-lookup"><span data-stu-id="03d09-239">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="03d09-240">**28 novembre 2017 :** mise à disposition\* de la prise en charge de l’annulation sur des fonctions asynchrones (nécessite la modification des fonctions de flux)</span><span class="sxs-lookup"><span data-stu-id="03d09-240">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="03d09-241">\*\* 7 mai 2018 : support\* fourni pour Mac, Excel Online et fonctions synchrones en cours de \*\*traitement</span><span class="sxs-lookup"><span data-stu-id="03d09-241">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="03d09-242">**20 septembre 2018** : Support fourni pour les fonctions personnalisées à l'exécution de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="03d09-242">**September 20, 2018**: Shipped support for custom functions JavaScript runtime.</span></span> <span data-ttu-id="03d09-243">Pour plus d’informations, voir [Exécution des fonctions personnalisées d’Excel](custom-functions-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="03d09-243">For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>

<span data-ttu-id="03d09-244">\* vers le canal Office Insiders</span><span class="sxs-lookup"><span data-stu-id="03d09-244">\* to the Office Insiders Channel</span></span>

## <a name="see-also"></a><span data-ttu-id="03d09-245">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="03d09-245">See also</span></span>

* [<span data-ttu-id="03d09-246">Métadonnées des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="03d09-246">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="03d09-247">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="03d09-247">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="03d09-248">Meilleures pratiques pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="03d09-248">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="03d09-249">Didacticiel sur les fonctions personnalisées d’Excel</span><span class="sxs-lookup"><span data-stu-id="03d09-249">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)