---
ms.date: 10/17/2018
description: Créer des fonctions personnalisées dans Excel à l’aide de JavaScript.
title: Créer des fonctions personnalisées dans Excel (Aperçu)
ms.openlocfilehash: 8383b5f6d568a1ce2da036fbacfb90404bbe8297
ms.sourcegitcommit: 2ac7d64bb2db75ace516a604866850fce5cb2174
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/14/2018
ms.locfileid: "26298550"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="8708c-103">Créer des fonctions personnalisées dans Excel (aperçu)</span><span class="sxs-lookup"><span data-stu-id="8708c-103">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="8708c-104">Les fonctions personnalisées permettent aux développeurs d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément.</span><span class="sxs-lookup"><span data-stu-id="8708c-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="8708c-105">Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="8708c-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="8708c-106">Cet article explique comment créer des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="8708c-106">This article explains how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="8708c-107">L’illustration suivante montre un utilisateur final insérant une fonction personnalisée dans une cellule de feuille de calcul Excel.</span><span class="sxs-lookup"><span data-stu-id="8708c-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="8708c-108">Le `CONTOSO.ADD42` fonction personnalisée est conçue pour ajouter 42 à la paire de nombres que spécifie l’utilisateur en tant que paramètres d’entrée de la fonction.</span><span class="sxs-lookup"><span data-stu-id="8708c-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="8708c-109">Le code suivant définit la `ADD42` fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="8708c-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="8708c-110">La section [problèmes connus](#known-issues)plus loin dans cet article indique les limitations en cours de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="8708c-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="8708c-111">Composants d’un projet de complément fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="8708c-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="8708c-112">Si vous utilisez le [générateur Yo Office](https://github.com/OfficeDev/generator-office) pour créer un projet complément de fonctions personnalisées Excel, vous verrez les fichiers suivants dans le projet crée par le générateur :</span><span class="sxs-lookup"><span data-stu-id="8708c-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="8708c-113">Fichier</span><span class="sxs-lookup"><span data-stu-id="8708c-113">File</span></span> | <span data-ttu-id="8708c-114">Format de fichier</span><span class="sxs-lookup"><span data-stu-id="8708c-114">File format</span></span> | <span data-ttu-id="8708c-115">Description</span><span class="sxs-lookup"><span data-stu-id="8708c-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="8708c-116">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="8708c-116">**./src/customfunctions.js**</span></span><br/><span data-ttu-id="8708c-117">ou</span><span class="sxs-lookup"><span data-stu-id="8708c-117">or</span></span><br/><span data-ttu-id="8708c-118">**./src/customfunctions.ts**</span><span class="sxs-lookup"><span data-stu-id="8708c-118">**./src/customfunctions.ts**</span></span> | <span data-ttu-id="8708c-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="8708c-119">JavaScript</span></span><br/><span data-ttu-id="8708c-120">ou</span><span class="sxs-lookup"><span data-stu-id="8708c-120">or</span></span><br/><span data-ttu-id="8708c-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="8708c-121">TypeScript</span></span> | <span data-ttu-id="8708c-122">Contient le code qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="8708c-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="8708c-123">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="8708c-123">**./config/customfunctions.json**</span></span> | <span data-ttu-id="8708c-124">JSON</span><span class="sxs-lookup"><span data-stu-id="8708c-124">JSON</span></span> | <span data-ttu-id="8708c-125">Contient les métadonnées qui décrivent les fonctions personnalisées et permettent à Excel d’enregistrer les fonctions personnalisées et les rendre accessibles aux utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="8708c-125">Contains metadata that describes custom functions and enables Excel to register the custom functions and make them available to end users.</span></span> |
| <span data-ttu-id="8708c-126">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="8708c-126">**./index.html**</span></span> | <span data-ttu-id="8708c-127">HTML</span><span class="sxs-lookup"><span data-stu-id="8708c-127">HTML</span></span> | <span data-ttu-id="8708c-128">Fournit une référence&lt;script&gt; au fichier JavaScript qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="8708c-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="8708c-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="8708c-129">**Manifest.xml**</span></span> | <span data-ttu-id="8708c-130">XML</span><span class="sxs-lookup"><span data-stu-id="8708c-130">XML</span></span> | <span data-ttu-id="8708c-131">Spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers HTML, JavaScript et JSON qui figurent précédemment dans ce tableau.</span><span class="sxs-lookup"><span data-stu-id="8708c-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="8708c-132">Les sections suivantes vous apportent plus d'informations sur ces fichiers.</span><span class="sxs-lookup"><span data-stu-id="8708c-132">The following sections provide more information about those changes.</span></span>

### <a name="script-file"></a><span data-ttu-id="8708c-133">Fichier de script</span><span class="sxs-lookup"><span data-stu-id="8708c-133">Script file</span></span> 

<span data-ttu-id="8708c-134">Le fichier de script (**./src/customfunctions.js** ou **./src/customfunctions.ts** du projet créé par le Générateur de Yo Office) contient le code qui définit les fonctions personnalisées et mappe les noms des fonctions personnalisées aux objets dans le [fichier de métadonnées JSON](#json-metadata-file).</span><span class="sxs-lookup"><span data-stu-id="8708c-134">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="8708c-135">Par exemple, le code suivant définit les fonctions personnalisées `add` et `increment`indique ensuite les informations de mappage pour les deux fonctions.</span><span class="sxs-lookup"><span data-stu-id="8708c-135">For example, the following code defines the custom functions `add` and `increment` and then specifies mapping information for both functions.</span></span> <span data-ttu-id="8708c-136">La fonction `add` mappée à l’objet dans le fichier de métadonnées JSON où la valeur de la `id` propriété est **AJOUTER**et la fonction`increment`mappée à l’objet dans le fichier de métadonnées dans laquelle la valeur de la `id` propriété est **INCRÉMENT**.</span><span class="sxs-lookup"><span data-stu-id="8708c-136">The `add` function is mapped to the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is mapped to the object in the metadata file where the value of the `id` property is **INCREMENT**.</span></span> <span data-ttu-id="8708c-137">Voir [Recommandations fonctions personnalisées](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) pour plus d’informations sur le mappage des noms de fonction dans le fichier de script pour objets dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="8708c-137">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about mapping function names in the script file to objects in the JSON metadata file.</span></span>

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

### <a name="json-metadata-file"></a><span data-ttu-id="8708c-138">Fichier de métadonnées JSON</span><span class="sxs-lookup"><span data-stu-id="8708c-138">JSON metadata file</span></span> 

<span data-ttu-id="8708c-139">Le fichier de métadonnées fonctions personnalisées (**./config/customfunctions.json** du projet créé par le Générateur de Yo Office) fournit les informations dont Excel a besoin pour enregistrer les fonctions personnalisées et les rendre disponibles aux utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="8708c-139">The custom functions metadata file (**./config/customfunctions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users.</span></span> <span data-ttu-id="8708c-140">Les fonctions personnalisées sont enregistrées lorsqu’un utilisateur lance un complément pour la première fois.</span><span class="sxs-lookup"><span data-stu-id="8708c-140">Custom functions are registered when a user runs an add-in for the first time.</span></span> <span data-ttu-id="8708c-141">Après cela, elles sont disponibles pour cet utilisateur depuis tous les classeurs (c'est-à-dire pas seulement dans le classeur dans lequel le complément est initialement exécuté.)</span><span class="sxs-lookup"><span data-stu-id="8708c-141">After that, they are available to that same user in all workbooks (i.e., not only in the workbook where the add-in initially ran.)</span></span>

> [!TIP]
> <span data-ttu-id="8708c-142">Les paramètres du serveur qui héberge le fichier JSON doivent avoir [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) activée afin que les fonctions personnalisées s’exécutent correctement dans Excel Online.</span><span class="sxs-lookup"><span data-stu-id="8708c-142">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="8708c-143">Le code suivant de **customfunctions.json** spécifie les métadonnées pour la `add` fonction et la `increment` fonction qui ont été décrites précédemment.</span><span class="sxs-lookup"><span data-stu-id="8708c-143">The following code in **customfunctions.json** specifies the metadata for the `add` function and the `increment` function that were described previously.</span></span> <span data-ttu-id="8708c-144">Le tableau qui suit cet exemple de code fournit des informations détaillées sur les propriétés individuelles au sein de cet objet JSON.</span><span class="sxs-lookup"><span data-stu-id="8708c-144">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span> <span data-ttu-id="8708c-145">Voir [Recommandations fonctions personnalisées](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) pour plus d’informations sur la spécification de la valeur de `id` et les propriétés`name`dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="8708c-145">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

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

<span data-ttu-id="8708c-146">Le tableau suivant répertorie les propriétés généralement présentes dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="8708c-146">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="8708c-147">Pour plus d’informations sur le fichier de métadonnées JSON, voir [métadonnées fonctions personnalisées](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="8708c-147">For more detailed information about the JSON metadata file, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="8708c-148">Propriété</span><span class="sxs-lookup"><span data-stu-id="8708c-148">Property</span></span>  | <span data-ttu-id="8708c-149">Description</span><span class="sxs-lookup"><span data-stu-id="8708c-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="8708c-150">Un ID unique pour la fonction.</span><span class="sxs-lookup"><span data-stu-id="8708c-150">A unique ID for the group.</span></span> <span data-ttu-id="8708c-151">Cet ID peut contenir uniquement des points et caractères alphanumériques et ne doit pas être modifié une fois défini.</span><span class="sxs-lookup"><span data-stu-id="8708c-151">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="8708c-152">Nom de la fonction que voit l’utilisateur final dans Excel.</span><span class="sxs-lookup"><span data-stu-id="8708c-152">Name of the function that the end user sees in Excel.</span></span> <span data-ttu-id="8708c-153">Dans Excel, ce nom de la fonction sera précédé de l’espace de noms de fonctions personnalisées spécifié dans le [fichier manifeste XML](#manifest-file).</span><span class="sxs-lookup"><span data-stu-id="8708c-153">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the [XML manifest file](#manifest-file).</span></span> |
| `helpUrl` | <span data-ttu-id="8708c-154">URL de la page qui s’affiche quand un utilisateur demande de l’aide.</span><span class="sxs-lookup"><span data-stu-id="8708c-154">URL for the page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="8708c-155">Descriptif de la fonction.</span><span class="sxs-lookup"><span data-stu-id="8708c-155">Describes what the function does.</span></span> <span data-ttu-id="8708c-156">Cette valeur apparaît comme une info-bulle lorsque la fonction est l’élément sélectionné dans le menu de saisie semi-automatique des formules dans Excel.</span><span class="sxs-lookup"><span data-stu-id="8708c-156">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="8708c-157">Objet qui définit le type d’informations renvoyées par la fonction.</span><span class="sxs-lookup"><span data-stu-id="8708c-157">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="8708c-158">Pour plus d’informations sur cet objet, voir [résultat](custom-functions-json.md#result).</span><span class="sxs-lookup"><span data-stu-id="8708c-158">For detailed information about this object, see [result](custom-functions-json.md#result).</span></span> |
| `parameters` | <span data-ttu-id="8708c-159">Tableau qui définit les paramètres d’entrée de la fonction.</span><span class="sxs-lookup"><span data-stu-id="8708c-159">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="8708c-160">Pour plus d’informations sur cet objet, voir [paramètres](custom-functions-json.md#parameters).</span><span class="sxs-lookup"><span data-stu-id="8708c-160">For detailed information about this object, see [parameters](custom-functions-json.md#parameters).</span></span> |
| `options` | <span data-ttu-id="8708c-161">Vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction.</span><span class="sxs-lookup"><span data-stu-id="8708c-161">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="8708c-162">Pour plus d’informations sur l’utilisation de cette propriété, voir [Diffusion en continu de fonctions](#streaming-functions) et [Annuler une fonction](#canceling-a-function) plus loin dans cet article.</span><span class="sxs-lookup"><span data-stu-id="8708c-162">For more information about how this property can be used, see [Streaming functions](#streaming-functions) and [Canceling a function](#canceling-a-function) later in this article.</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="8708c-163">Fichier manifeste</span><span class="sxs-lookup"><span data-stu-id="8708c-163">Manifest file</span></span>

<span data-ttu-id="8708c-164">Le fichier manifeste XML pour un complément qui définit les fonctions personnalisées (**./manifest.xml** du projet créé par le Générateur de Yo Office) spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers HTML, JavaScript et JSON.</span><span class="sxs-lookup"><span data-stu-id="8708c-164">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="8708c-165">Le balisage XML suivant montre un exemple des éléments`<ExtensionPoint>` et `<Resources>` que vous devez inclure dans manifeste d’un complément pour activer les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="8708c-165">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span>  

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
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. Can only contain alphanumeric characters and periods.-->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="8708c-166">Les fonctions dans Excel sont précédées par l’espace de noms spécifié dans votre fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="8708c-166">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="8708c-167">L’espace de noms d’une fonction vient avant le nom de fonction et les deux sont séparés par un point.</span><span class="sxs-lookup"><span data-stu-id="8708c-167">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="8708c-168">Par exemple, pour appeler la fonction `ADD42` dans la cellule de feuille de calcul Excel, vous saisiriez `=CONTOSO.ADD42`, car `CONTOSO` est l’espace de noms et `ADD42` est le nom de la fonction spécifié dans le fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="8708c-168">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="8708c-169">L’espace de noms est destiné à être utilisé comme identificateur de votre entreprise ou du complément.</span><span class="sxs-lookup"><span data-stu-id="8708c-169">The prefix is intended to be used as an identifier for your add-in.</span></span> <span data-ttu-id="8708c-170">Un espace de noms ne peut contenir que des points et des caractères alphanumériques.</span><span class="sxs-lookup"><span data-stu-id="8708c-170">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="8708c-171">Fonctions qui retournent des données provenant de sources externes</span><span class="sxs-lookup"><span data-stu-id="8708c-171">Functions that return data from external sources</span></span>

<span data-ttu-id="8708c-172">Si une fonction personnalisée récupère des données d’une source externe comme le web, elle doit :</span><span class="sxs-lookup"><span data-stu-id="8708c-172">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="8708c-173">Renvoyer une promesse JavaScript à Excel.</span><span class="sxs-lookup"><span data-stu-id="8708c-173">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="8708c-174">Résoudre la promesse avec la valeur finale à l’aide de la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="8708c-174">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="8708c-175">Les fonctions personnalisées affichent un `#GETTING_DATA`résultat temporaire dans la cellule, tandis qu’ Excel attend que le résultat final.</span><span class="sxs-lookup"><span data-stu-id="8708c-175">Custom functions display a `#GETTING_DATA` temporary result in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="8708c-176">Les utilisateurs peuvent interagir normalement avec le reste de la feuille de calcul pendant qu’ils attendent le résultat.</span><span class="sxs-lookup"><span data-stu-id="8708c-176">Users can interact normally with the rest of the worksheet while they wait for the result.</span></span>

<span data-ttu-id="8708c-177">Le code suivant indique un exemple de`getTemperature()`fonction personnalisée qui récupère la température d’un thermomètre.</span><span class="sxs-lookup"><span data-stu-id="8708c-177">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer.</span></span> <span data-ttu-id="8708c-178">Notez que `sendWebRequest` est une fonction hypothétique (non spécifiée ici) qui utilise [XHR](custom-functions-runtime.md#xhr-example) pour appeler un service web de température.</span><span class="sxs-lookup"><span data-stu-id="8708c-178">Note that `sendWebRequest` is a hypothetical function (not specified here) that uses [XHR](custom-functions-runtime.md#xhr-example) to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streaming-functions"></a><span data-ttu-id="8708c-179">Fonctions de diffusion en continu</span><span class="sxs-lookup"><span data-stu-id="8708c-179">Streaming functions</span></span>

<span data-ttu-id="8708c-180">Les fonctions personnalisées de diffusion en continu vous aident à copier des données à des cellules à plusieurs reprises au fil du temps, sans exiger qu’un utilisateur demande explicitement l’actualisation des données.</span><span class="sxs-lookup"><span data-stu-id="8708c-180">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span> <span data-ttu-id="8708c-181">L’exemple de code suivant est une fonction personnalisée qui ajoute un nombre au résultat chaque seconde.</span><span class="sxs-lookup"><span data-stu-id="8708c-181">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="8708c-182">Tenez compte des informations suivantes à propos de ce code :</span><span class="sxs-lookup"><span data-stu-id="8708c-182">Note the following about this code:</span></span>

- <span data-ttu-id="8708c-183">Excel affiche chaque nouvelle valeur automatiquement à l’aide du `setResult` rappel.</span><span class="sxs-lookup"><span data-stu-id="8708c-183">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="8708c-184">Le deuxième paramètre d’entrée `handler`, n’est pas visible aux utilisateurs finaux dans Excel lorsqu’ils sélectionnent la fonction à partir du menu de saisie semi-automatique.</span><span class="sxs-lookup"><span data-stu-id="8708c-184">The second input parameter, `handler`, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>

- <span data-ttu-id="8708c-185">Le `onCanceled` rappel définit la fonction qui s’exécute lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="8708c-185">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="8708c-186">Vous devez implémenter un gestionnaire d’annulation comme suit pour n’importe quelle fonction de diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="8708c-186">You must implement a cancellation handler like this for any streaming function.</span></span> <span data-ttu-id="8708c-187">Pour plus d’informations, voir [Annuler une fonction](#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="8708c-187">For more information, see [Canceling a function](#canceling-a-function).</span></span>

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

<span data-ttu-id="8708c-188">Lorsque vous spécifiez des métadonnées pour une fonction de diffusion en continu dans le fichier de métadonnées JSON, vous devez définir les propriétés `"cancelable": true` et `"stream": true` au sein de l’objet`options`, comme illustré dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="8708c-188">When you specify metadata for a streaming function in the JSON metadata file, you must set the properties `"cancelable": true` and `"stream": true` within the `options` object, as shown in the following example.</span></span>

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

## <a name="canceling-a-function"></a><span data-ttu-id="8708c-189">Annulation d’une fonction</span><span class="sxs-lookup"><span data-stu-id="8708c-189">Canceling a function</span></span>

<span data-ttu-id="8708c-190">Dans certains cas, vous devrez annuler l’exécution d’une fonction personnalisée de diffusion en continu pour réduire la consommation de bande passante, de la mémoire de travail et la charge du CPU.</span><span class="sxs-lookup"><span data-stu-id="8708c-190">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="8708c-191">Excel annule l’exécution d’une fonction dans les situations suivantes :</span><span class="sxs-lookup"><span data-stu-id="8708c-191">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="8708c-192">L’utilisateur modifie ou supprime une cellule qui fait référence à la fonction.</span><span class="sxs-lookup"><span data-stu-id="8708c-192">The user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="8708c-193">Un des arguments (entrées) de la fonction est modifié.</span><span class="sxs-lookup"><span data-stu-id="8708c-193">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="8708c-194">Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.</span><span class="sxs-lookup"><span data-stu-id="8708c-194">In this case, a new function call is triggered in addition to the cancelation.</span></span>

- <span data-ttu-id="8708c-195">L’utilisateur déclenche manuellement le recalcul.</span><span class="sxs-lookup"><span data-stu-id="8708c-195">When the user triggers recalculation manually.</span></span> <span data-ttu-id="8708c-196">Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.</span><span class="sxs-lookup"><span data-stu-id="8708c-196">In this case, a new function call is triggered in addition to the cancelation.</span></span>

<span data-ttu-id="8708c-197">Pour activer la possibilité d’annuler une fonction, vous devez implémenter un gestionnaire d’annulation au sein de la fonction JavaScript et spécifier la propriété `"cancelable": true` au sein de l’objet`options` dans les métadonnées JSON décrivant la fonction.</span><span class="sxs-lookup"><span data-stu-id="8708c-197">To enable the ability to cancel a function, you must implement a cancellation handler within the JavaScript function and specify the property `"cancelable": true` within the `options` object in the JSON metadata that describes the function.</span></span> <span data-ttu-id="8708c-198">Les exemples de code dans la section précédente de cet article fournissent un exemple de ces techniques.</span><span class="sxs-lookup"><span data-stu-id="8708c-198">The code samples in the previous section of this article provide an example of these techniques.</span></span>

## <a name="saving-and-sharing-state"></a><span data-ttu-id="8708c-199">Enregistrement et partage d’état</span><span class="sxs-lookup"><span data-stu-id="8708c-199">Saving and sharing state</span></span>

<span data-ttu-id="8708c-200">Les fonctions personnalisées peuvent enregistrer des données dans des variables JavaScript globales, qui peuvent être utilisées dans les appels suivants.</span><span class="sxs-lookup"><span data-stu-id="8708c-200">Custom functions can save data in global JavaScript variables, which can be used in subsequent calls.</span></span> <span data-ttu-id="8708c-201">Un état enregistré est utile lorsque les utilisateurs appellent la même fonction personnalisée à partir de plusieurs cellules, car toutes les instances de la fonction pouvant accéder à l’état.</span><span class="sxs-lookup"><span data-stu-id="8708c-201">Saved state is useful when users call the same custom function from more than one cell, because all instances of the function can access the state.</span></span> <span data-ttu-id="8708c-202">Par exemple, vous pouvez enregistrer les données renvoyées par un appel à une ressource web pour éviter d’effectuer des appels supplémentaires à la même ressource web.</span><span class="sxs-lookup"><span data-stu-id="8708c-202">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="8708c-203">L’exemple de code suivant montre une implémentation d’une fonction de diffusion en continu de la température qui enregistre l’état global.</span><span class="sxs-lookup"><span data-stu-id="8708c-203">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="8708c-204">Tenez compte des informations suivantes à propos de ce code :</span><span class="sxs-lookup"><span data-stu-id="8708c-204">Note the following about this code:</span></span>

- <span data-ttu-id="8708c-205">La fonction`streamTemperature`met à jour la valeur de température qui s’affiche dans la cellule chaque seconde et elle utilise la `savedTemperatures` variable en tant que source de données.</span><span class="sxs-lookup"><span data-stu-id="8708c-205">The `streamTemperature` function updates the temperature value that's displayed in the cell every second and it uses the `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="8708c-206">Étant donné que `streamTemperature` est une fonction de diffusion en continu, elle implémente un gestionnaire d’annulation qui s’exécute lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="8708c-206">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="8708c-207">Si un utilisateur appelle la `streamTemperature` fonction à partir de plusieurs cellules dans Excel, la`streamTemperature` fonction lit les données dans la même `savedTemperatures` variable à chaque fois qu’elle s’exécute.</span><span class="sxs-lookup"><span data-stu-id="8708c-207">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="8708c-208">La `refreshTemperature` fonction lit la température d’un thermomètre spécifique à chaque seconde qui passe et stocke le résultat dans la`savedTemperatures`variable.</span><span class="sxs-lookup"><span data-stu-id="8708c-208">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="8708c-209">Étant donné que la `refreshTemperature` fonction n’est pas exposée aux utilisateurs finaux dans Excel, elle ne doit pas être enregistrées dans le fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="8708c-209">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
  if(!savedTemperatures[thermometerID]){
    refreshTemperature(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
  }

  function getNextTemperature(){
    handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
    var delayTime = 1000; // Amount of milliseconds to delay a request by.
    setTimeout(getNextTemperature, delayTime); // Wait 1 second before updating Excel again.

    handler.onCancelled() = function {
      clearTimeout(delayTime);
    }
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

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="8708c-210">Utilisation des plages de données</span><span class="sxs-lookup"><span data-stu-id="8708c-210">Working with ranges of data</span></span>

<span data-ttu-id="8708c-211">Votre fonction personnalisée peut accepter une plage de données sous la forme d’un paramètre d’entrée, ou il peut renvoyer une plage de données.</span><span class="sxs-lookup"><span data-stu-id="8708c-211">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="8708c-212">Dans JavaScript, une plage de données est représentée sous la forme d’une matrice 2 dimensions.</span><span class="sxs-lookup"><span data-stu-id="8708c-212">In JavaScript, a range of data is represented as a 2-dimensional array.</span></span>

<span data-ttu-id="8708c-213">Par exemple, supposons que votre fonction renvoie la seconde valeur la plus élevée à partir d’une plage de nombres stockés dans Excel.</span><span class="sxs-lookup"><span data-stu-id="8708c-213">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="8708c-214">La fonction suivante prend le paramètre `values`, c’est-à-dire un type de `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="8708c-214">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="8708c-215">Notez que dans les métadonnées JSON pour cette fonction, vous devez définir la propriété `type` de paramètre à `matrix`.</span><span class="sxs-lookup"><span data-stu-id="8708c-215">Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="handling-errors"></a><span data-ttu-id="8708c-216">Gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="8708c-216">Handling errors</span></span>

<span data-ttu-id="8708c-217">Lorsque vous créez un complément à l’aide des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution.</span><span class="sxs-lookup"><span data-stu-id="8708c-217">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="8708c-218">La gestion des erreurs pour fonctions personnalisées est identique à la[gestion des erreurs pour l’API JavaScript Excel](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="8708c-218">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="8708c-219">Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.</span><span class="sxs-lookup"><span data-stu-id="8708c-219">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="known-issues"></a><span data-ttu-id="8708c-220">Problèmes connus</span><span class="sxs-lookup"><span data-stu-id="8708c-220">Known issues</span></span>

- <span data-ttu-id="8708c-221">Les descriptions de paramètre et les URL d’aide ne sont pas encore utilisés par Excel.</span><span class="sxs-lookup"><span data-stu-id="8708c-221">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="8708c-222">Les fonctions personnalisées ne sont actuellement pas disponibles dans Excel pour les clients mobiles.</span><span class="sxs-lookup"><span data-stu-id="8708c-222">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="8708c-223">Les fonctions volatiles (celles qui sont recalculées à chaque fois que des données autonomes sont modifiées dans la feuille de calcul) ne sont pas encore prises en charge.</span><span class="sxs-lookup"><span data-stu-id="8708c-223">Volatile functions (those that recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="8708c-224">Le déploiement via le portail d’administration Office 365 et AppSource n’est pas encore activé.</span><span class="sxs-lookup"><span data-stu-id="8708c-224">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="8708c-225">Les fonctions personnalisées dans Excel Online peuvent cesser de fonctionner pendant une session après une période d’inactivité.</span><span class="sxs-lookup"><span data-stu-id="8708c-225">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="8708c-226">Actualiser la page du navigateur (F5), puis entrez une fonction personnalisée pour restaurer la fonctionnalité.</span><span class="sxs-lookup"><span data-stu-id="8708c-226">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="8708c-227">Vous pouvez voir le **## CHARGEMENT_DONNEES** résultat temporaire au sein des cellules d’une feuille de calcul si vous avez plusieurs compléments en cours d’exécution sur Excel pour Windows.</span><span class="sxs-lookup"><span data-stu-id="8708c-227">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows.</span></span> <span data-ttu-id="8708c-228">Fermez toutes les fenêtres Excel et redémarrez Excel.</span><span class="sxs-lookup"><span data-stu-id="8708c-228">Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="8708c-229">Des outils de débogage spécifiques aux fonctions personnalisées seront peut-être disponibles à l’avenir.</span><span class="sxs-lookup"><span data-stu-id="8708c-229">Debugging tools specifically for custom functions may be available in the future.</span></span> <span data-ttu-id="8708c-230">En attendant, vous pouvez déboguer sur Excel Online à l’aide des outils de développement F12.</span><span class="sxs-lookup"><span data-stu-id="8708c-230">In the meantime, you can debug on Excel Online using F12 developer tools.</span></span> <span data-ttu-id="8708c-231">Plus de détails dans [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="8708c-231">See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>

## <a name="changelog"></a><span data-ttu-id="8708c-232">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="8708c-232">Changelog</span></span>

- <span data-ttu-id="8708c-233">**7 novembre 2017 :** mise à disposition des exemples et de l’aperçu des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="8708c-233">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="8708c-234">**20 novembre 2017 :** correction du bogue de compatibilité pour les utilisateurs de la version 8801 et ultérieure</span><span class="sxs-lookup"><span data-stu-id="8708c-234">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="8708c-235">**28 novembre 2017 :** prise en charge de l’annulation sur des fonctions asynchrones (nécessite la modification des fonctions de flux)</span><span class="sxs-lookup"><span data-stu-id="8708c-235">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="8708c-236">**7 mai 2018**: prise en charge pour Mac, Excel Online et fonctions synchrones dans les processus en cours d’exécution</span><span class="sxs-lookup"><span data-stu-id="8708c-236">**May 7, 2018**: Shipped\* support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="8708c-237">**20 septembre 2018**: prise en charge de fonctions personnalisées lors de l’exécution de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="8708c-237">**September 20, 2018**: Shipped support for custom functions JavaScript runtime.</span></span> <span data-ttu-id="8708c-238">Pour plus d’informations, voir [Runtime pour les fonctions personnalisées Excel](custom-functions-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="8708c-238">For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>
- <span data-ttu-id="8708c-239">**20 octobre 2018**: avec le programme[October Insiders build](https://support.office.com/fr-FR/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24), les fonctions personnalisées nécessitent désormais le paramètre « id » dans votre[métadonnées fonctions personnalisées](custom-functions-json.md) pour les versions Windows Bureau et Online.</span><span class="sxs-lookup"><span data-stu-id="8708c-239">**October 20, 2018**: With the [October Insiders build](https://support.office.com/fr-FR/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24), Custom Functions now requires the 'id' parameter in your [custom functions metadata](custom-functions-json.md) for Windows Desktop and Online.</span></span> <span data-ttu-id="8708c-240">Sur Mac, ce paramètre doit être ignoré.</span><span class="sxs-lookup"><span data-stu-id="8708c-240">On Mac, this parameter should be ignored.</span></span>


<span data-ttu-id="8708c-241">\* pour la chaîne [Office Insider](https://products.office.com/office-insider) (anciennement appelée « Insider Fast »)</span><span class="sxs-lookup"><span data-stu-id="8708c-241">\* to the [Office Insider](https://products.office.com/office-insider) channel (formerly called "Insider Fast")</span></span>

## <a name="see-also"></a><span data-ttu-id="8708c-242">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8708c-242">See also</span></span>

* [<span data-ttu-id="8708c-243">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="8708c-243">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="8708c-244">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="8708c-244">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="8708c-245">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="8708c-245">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="8708c-246">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="8708c-246">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
