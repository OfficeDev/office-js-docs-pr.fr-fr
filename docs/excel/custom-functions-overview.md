---
ms.date: 10/09/2018
description: Créer des fonctions personnalisées dans Excel à l’aide de JavaScript.
title: Créer des fonctions personnalisées dans Excel (Aperçu)
ms.openlocfilehash: e52039f2618f793f688cd89c5d62bac0a8632667
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506118"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="60ad0-103">Créer des fonctions personnalisées dans Excel (aperçu)</span><span class="sxs-lookup"><span data-stu-id="60ad0-103">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="60ad0-p101">Les fonctions personnalisées permettent aux développeurs d'ajouter de nouvelles fonctions à Excel en définissant ces fonctions dans JavaScript comme partie d’un complément. Les utilisateurs d'Excel peuvent accéder à des fonctions personnalisées comme n'importe quelle fonction native d'Excel, telle que `SUM()`. Cet article explique comment créer des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="60ad0-p101">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`. This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="60ad0-p102">L’illustration suivante montre un utilisateur final insérant une fonction personnalisée dans une cellule de feuille de calcul Excel. La fonction personnalisée  `CONTOSO.ADD42` est conçue pour ajouter 42 à la paire de nombres spécifiée par l’utilisateur en tant que paramètres d’entrée à la fonction.</span><span class="sxs-lookup"><span data-stu-id="60ad0-p102">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet. The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="60ad0-109">Le code suivant définit la fonction personnalisée `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="60ad0-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="60ad0-110">Plus loin dans cet article, la section [Problèmes connus](#known-issues) indique les limites actuelles des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="60ad0-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="60ad0-111">Composants d’un projet de complément de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="60ad0-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="60ad0-112">Si vous utilisez le [générateur de Yo Office](https://github.com/OfficeDev/generator-office) pour créer un projet de complément de fonctions personnalisées Excel, vous verrez les fichiers suivants dans le projet que le générateur crée :</span><span class="sxs-lookup"><span data-stu-id="60ad0-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="60ad0-113">Fichier</span><span class="sxs-lookup"><span data-stu-id="60ad0-113">File</span></span> | <span data-ttu-id="60ad0-114">Format de fichier</span><span class="sxs-lookup"><span data-stu-id="60ad0-114">File format</span></span> | <span data-ttu-id="60ad0-115">Description</span><span class="sxs-lookup"><span data-stu-id="60ad0-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="60ad0-116">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="60ad0-116">**./src/customfunctions.js**</span></span><br/><span data-ttu-id="60ad0-117">ou</span><span class="sxs-lookup"><span data-stu-id="60ad0-117">or</span></span><br/><span data-ttu-id="60ad0-118">**./src/customfunctions.ts**</span><span class="sxs-lookup"><span data-stu-id="60ad0-118">**./src/customfunctions.ts**</span></span> | <span data-ttu-id="60ad0-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="60ad0-119">JavaScript</span></span><br/><span data-ttu-id="60ad0-120">ou</span><span class="sxs-lookup"><span data-stu-id="60ad0-120">or</span></span><br/><span data-ttu-id="60ad0-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="60ad0-121">TypeScript</span></span> | <span data-ttu-id="60ad0-122">Contient le code qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="60ad0-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="60ad0-123">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="60ad0-123">**./config/customfunctions.json**</span></span> | <span data-ttu-id="60ad0-124">JSON</span><span class="sxs-lookup"><span data-stu-id="60ad0-124">JSON</span></span> | <span data-ttu-id="60ad0-125">Contient des métadonnées qui décrivent les fonctions personnalisées et permettent à Excel d'enregistrer les fonctions personnalisées et de les mettre à la disposition des utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="60ad0-125">Contains metadata that describes custom functions and enables Excel to register the custom functions in order to make them available to end-users.</span></span> |
| <span data-ttu-id="60ad0-126">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="60ad0-126">**./index.html**</span></span> | <span data-ttu-id="60ad0-127">HTML</span><span class="sxs-lookup"><span data-stu-id="60ad0-127">HTML</span></span> | <span data-ttu-id="60ad0-128">Fournit une référence de &lt;script&gt; pour le fichier JavaScript qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="60ad0-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="60ad0-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="60ad0-129">**Manifest.xml**</span></span> | <span data-ttu-id="60ad0-130">XML</span><span class="sxs-lookup"><span data-stu-id="60ad0-130">XML</span></span> | <span data-ttu-id="60ad0-131">Spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers JavaScript, JSON et HTML répertoriés précédemment dans ce tableau.</span><span class="sxs-lookup"><span data-stu-id="60ad0-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="60ad0-132">Les sections suivantes fournissent plus d’informations sur ces fichiers.</span><span class="sxs-lookup"><span data-stu-id="60ad0-132">The following sections provide more information about those changes.</span></span>

### <a name="script-file"></a><span data-ttu-id="60ad0-133">Fichier de script</span><span class="sxs-lookup"><span data-stu-id="60ad0-133">Script file</span></span> 

<span data-ttu-id="60ad0-134">Le fichier de script (**./src/customfunctions.js** ou **./src/customfunctions.ts** dans le projet que le générateur de Yo Office crée) contient le code qui définit les fonctions personnalisées et mappe les noms des fonctions personnalisées aux objets du [fichier de métadonnées JSON](#json-metadata-file).</span><span class="sxs-lookup"><span data-stu-id="60ad0-134">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="60ad0-p103">Par exemple, le code suivant définit les fonctions personnalisées `add` et `increment` , puis spécifie les informations de mappage pour les deux fonctions. La fonction  `add` est mappée à l’objet dans le fichier de métadonnées JSON où la valeur de la propriété  `id` est **ADD**et la fonction  `increment` est mappée à l’objet dans le fichier de métadonnées où la valeur de la propriété `id` est **INCREMENT**. Pour plus d’informations sur le mappage de noms de fonction dans le fichier de script à des objets dans le fichier de métadonnées JSON, consultez les [meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) .</span><span class="sxs-lookup"><span data-stu-id="60ad0-p103">For example, the following code defines the custom functions `add` and `increment` and then specifies mapping information for both functions. The `add` function is mapped to the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is mapped to the object in the metadata file where the value of the `id` property is **INCREMENT**. See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about mapping function names in the script file to objects in the JSON metadata file.</span></span>

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

### <a name="json-metadata-file"></a><span data-ttu-id="60ad0-138">Fichier de métadonnées JSON</span><span class="sxs-lookup"><span data-stu-id="60ad0-138">JSON metadata file</span></span> 

<span data-ttu-id="60ad0-p104">Le fichier de métadonnées des fonctions personnalisées (**./config/customfunctions.json** dans le projet que crée le générateur de Office Yo) fournit les informations nécessaires à Excel pour enregistrer des fonctions personnalisées et les rendre disponibles pour les utilisateurs finaux. Les fonctions personnalisées sont enregistrées lorsqu’un utilisateur exécute un complément pour la première fois. Après cela, elles sont disponibles pour cet utilisateur dans tous les classeurs (autrement dit, pas seulement dans le classeur dans lequel le complément a été exécuté initialement).</span><span class="sxs-lookup"><span data-stu-id="60ad0-p104">The custom functions metadata file (**./config/customfunctions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users. Custom functions are registered when a user runs an add-in for the first time. After that, they are available to that same user in all workbooks (i.e., not only in the workbook where the add-in initially ran.)</span></span>

> [!TIP]
> <span data-ttu-id="60ad0-142">Parmi les paramètres de serveur sur le serveur qui héberge le fichier JSON, [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) doit être activé pour que les fonctions personnalisées s'exécutent correctement dans Excel Online.</span><span class="sxs-lookup"><span data-stu-id="60ad0-142">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="60ad0-p105">Le code suivant dans le fichier **customfunctions.json** spécifie les métadonnées pour les fonctions `add` et `increment` précédemment décrites. Le tableau qui suit cet exemple de code fournit des informations détaillées sur les propriétés de cet objet JSON. Consultez les [meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) pour plus d’informations sur la spécification de la valeur des propriétés `id` et `name`  dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="60ad0-p105">The following code in **customfunctions.json** specifies the metadata for the `add` function and the `increment` function that were described previously. The table that follows this code sample provides detailed information about the individual properties within this JSON object. See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

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

<span data-ttu-id="60ad0-p106">Le tableau suivant répertorie les propriétés qui sont généralement présentes dans le fichier de métadonnées JSON. Pour plus d’informations sur le fichier de métadonnées JSON, voir [fonctions de métadonnées personnalisées](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="60ad0-p106">The following table lists the properties that are typically present in the JSON metadata file. For more detailed information about the JSON metadata file, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="60ad0-148">Propriété</span><span class="sxs-lookup"><span data-stu-id="60ad0-148">Property</span></span>  | <span data-ttu-id="60ad0-149">Description</span><span class="sxs-lookup"><span data-stu-id="60ad0-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="60ad0-p107">ID unique de la fonction. Cet ID ne doit pas être modifié après sa définition.</span><span class="sxs-lookup"><span data-stu-id="60ad0-p107">A unique ID for the function. This ID should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="60ad0-p108">Nom de la fonction que l’utilisateur final voit dans Excel. Dans Excel, ce nom de fonction aura pour préfixe l’espace de noms des fonctions personnalisées qui est spécifié dans le [fichier manifeste XML](#manifest-file).</span><span class="sxs-lookup"><span data-stu-id="60ad0-p108">Name of the function that the end user sees in Excel. In Excel, this function name will be prefixed by the custom functions namespace that's specified in the [XML manifest file](#manifest-file).</span></span> |
| `helpUrl` | <span data-ttu-id="60ad0-154">URL de la page qui s’affiche lorsqu’un utilisateur demande de l’aide.</span><span class="sxs-lookup"><span data-stu-id="60ad0-154">Url for a page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="60ad0-p109">Décrit l'action de la fonction. Cette valeur s’affiche comme une info-bulle lorsque la fonction est l’élément sélectionné dans le menu de saisie semi-automatique dans Excel.</span><span class="sxs-lookup"><span data-stu-id="60ad0-p109">Describes what the function does. This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="60ad0-p110">Objet qui définit le type d’informations renvoyées par la fonction. La valeur de la propriété enfant `type` peut être **string**, **number** ou **boolean**. La valeur de la propriété enfant `dimensionality` peut être **scalar** ou **matrix** (un tableau à deux dimensions des valeurs du `type` spécifié).</span><span class="sxs-lookup"><span data-stu-id="60ad0-p110">Object that defines the type of information that is returned by the function. The value of the `type` child property can be **string**, **number**, or **boolean**. The value of the `dimensionality` child property can be **scalar** or **matrix** (a two-dimensional array of values of the specified `type`).</span></span> |
| `parameters` | <span data-ttu-id="60ad0-p111">Tableau qui définit les paramètres d’entrée de la fonction. Les propriétés enfant `name` et `description` s’affichent dans intelliSense d'Excel. La valeur de la propriété enfant `type` peut être **string**, **number** ou **boolean**. La valeur de la propriété enfant `dimensionality` peut être **scala** ou la **matrix** (un tableau à deux dimensions des valeurs du `type` spécifié).</span><span class="sxs-lookup"><span data-stu-id="60ad0-p111">Array that defines the input parameters for the function. The `name` and `description` child properties appear in the Excel intelliSense. The value of the `type` child property can be **string**, **number**, or **boolean**. The value of the `dimensionality` child property can be **scalar** or **matrix** (a two-dimensional array of values of the specified `type`).</span></span> |
| `options` | <span data-ttu-id="60ad0-164">Vous permet de personnaliser certains aspects de la façon dont Excel exécute la fonction et quand.</span><span class="sxs-lookup"><span data-stu-id="60ad0-164">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="60ad0-165">Pour plus d’informations sur l’utilisation de cette propriété, voir [Fonctions de flux](#streaming-functions) et [Annulation d'une fonction](#canceling-a-function), plus loin dans cet article.</span><span class="sxs-lookup"><span data-stu-id="60ad0-165">For more information about how this property can be used, see [Streaming functions](#streaming-functions) and [Canceling a function](#canceling-a-function) later in this article.</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="60ad0-166">Fichier manifeste</span><span class="sxs-lookup"><span data-stu-id="60ad0-166">Manifest file</span></span>

<span data-ttu-id="60ad0-p113">Le fichier manifeste XML pour un complément qui définit les fonctions personnalisées (**./manifest.xml** dans le projet que crée le générateur de Office Yo) spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers JavaScript, JSON et HTML. Le code XML suivant montre un exemple d'éléments `<ExtensionPoint>` et `<Resources>` que vous devez inclure dans un manifeste de complément pour activer les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="60ad0-p113">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files. The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span>  

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
> <span data-ttu-id="60ad0-p114">Les fonctions d’Excel sont précédées de l'espace de noms spécifié dans votre fichier manifeste XML. L'espace de noms d’une fonction précède le nom de la fonction et ils sont séparés par un point. Par exemple, pour appeler la fonction `ADD42` dans une cellule de feuille de calcul Excel, vous devez saisir `=CONTOSO.ADD42`, étant donné que CONTOSO est l’espace de noms et `ADD42` est le nom de la fonction spécifiée dans le fichier JSON. L’espace de noms est destiné à être utilisé en tant qu’identificateur pour votre entreprise ou le complément.</span><span class="sxs-lookup"><span data-stu-id="60ad0-p114">Functions in Excel are prepended by the namespace specified in your XML manifest file. A function's namespace comes before the function name and they are separated by a period. For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because CONTOSO is the namespace and `ADD42` is the name of the function specified in the JSON file. The namespace is intended to be used as an identifier for your company or the add-in.</span></span> 

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="60ad0-173">Fonctions qui retournent des données provenant de sources externes</span><span class="sxs-lookup"><span data-stu-id="60ad0-173">Functions that return data from external sources</span></span>

<span data-ttu-id="60ad0-174">Si une fonction personnalisée récupère les données d’une source externe comme le Web, elle doit :</span><span class="sxs-lookup"><span data-stu-id="60ad0-174">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="60ad0-175">Renvoyer une promesse JavaScript à Excel.</span><span class="sxs-lookup"><span data-stu-id="60ad0-175">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="60ad0-176">Résolvez la promesse avec la valeur finale en utilisant la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="60ad0-176">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="60ad0-p115">Les fonctions personnalisées affichent un résultat temporaire `#GETTING_DATA` dans la cellule pendant qu’Excel attend le résultat final. Les utilisateurs peuvent interagir normalement avec le reste de la feuille de calcul pendant qu’ils attendent le résultat.</span><span class="sxs-lookup"><span data-stu-id="60ad0-p115">Custom functions display a `#GETTING_DATA` temporary result in the cell while Excel waits for the final result. Users can interact normally with the rest of the worksheet while they wait for the result.</span></span>

<span data-ttu-id="60ad0-p116">Dans l’exemple de code suivant, la fonction personnalisée  `getTemperature()` récupère la température actuelle d’un thermomètre. Notez que `sendWebRequest` est une fonction hypothétique  (non spécifiée ici) qui utilise [XHR](custom-functions-runtime.md#xhr-example) pour appeler un service web de température.</span><span class="sxs-lookup"><span data-stu-id="60ad0-p116">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer. Note that `sendWebRequest` is a hypothetical function (not specified here) that uses [XHR](custom-functions-runtime.md#xhr-example) to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streaming-functions"></a><span data-ttu-id="60ad0-181">Fonctions de diffusion en continu</span><span class="sxs-lookup"><span data-stu-id="60ad0-181">Streaming functions</span></span>

<span data-ttu-id="60ad0-182">Les fonctions de flux personnalisées vous permettent de transmettre des données aux cellules de manière répétée au fil du temps, sans qu'un utilisateur ait à demander explicitement une actualisation des données.</span><span class="sxs-lookup"><span data-stu-id="60ad0-182">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span> <span data-ttu-id="60ad0-183">L’échantillon de code suivant est une fonction personnalisée qui ajoute un nombre au résultat toutes les secondes.</span><span class="sxs-lookup"><span data-stu-id="60ad0-183">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="60ad0-184">Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="60ad0-184">Note the following about this code:</span></span>

- <span data-ttu-id="60ad0-185">Excel affiche automatiquement chaque nouvelle valeur en utilisant le rappel `setResult`.</span><span class="sxs-lookup"><span data-stu-id="60ad0-185">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="60ad0-186">Le second paramètre d’entrée, `handler`, n’est pas affiché pour les utilisateurs finaux dans Excel lorsqu’ils sélectionnent la fonction à partir du menu de saisie semi-automatique.</span><span class="sxs-lookup"><span data-stu-id="60ad0-186">The second input parameter, `handler`, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>

- <span data-ttu-id="60ad0-187">Le rappel `onCanceled` définit la fonction qui s’exécute lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="60ad0-187">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="60ad0-188">Vous devez implémenter un gestionnaire d'annulation comme celui-ci pour toute fonction de flux.</span><span class="sxs-lookup"><span data-stu-id="60ad0-188">You must implement a cancellation handler like this for any streaming function.</span></span> <span data-ttu-id="60ad0-189">Pour plus d’informations, voir [Annulation d’une fonction](#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="60ad0-189">For more information, see [Canceling a function](#canceling-a-function).</span></span>

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

<span data-ttu-id="60ad0-190">Lorsque vous spécifiez des métadonnées pour une fonction de diffusion en continu dans le fichier de métadonnées JSON, vous devez définir les propriétés `"cancelable": true` et `"stream": true` dans l'objet `options`, comme illustré dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="60ad0-190">When you specify metadata for a streamed function in the JSON metadata file, you must set the properties `"cancelable": true` and `"stream": true` within the `options` object, as shown in the following example.</span></span>

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

## <a name="canceling-a-function"></a><span data-ttu-id="60ad0-191">Annulation d’une fonction</span><span class="sxs-lookup"><span data-stu-id="60ad0-191">Canceling a function</span></span>

<span data-ttu-id="60ad0-192">Dans certains cas, vous devrez peut-être annuler l’exécution d’une fonction personnalisée en flux continu pour réduire la consommation de la bande passante, de la mémoire et de la charge processeur.</span><span class="sxs-lookup"><span data-stu-id="60ad0-192">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="60ad0-193">Excel annule l’exécution d’une fonction dans les situations suivantes :</span><span class="sxs-lookup"><span data-stu-id="60ad0-193">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="60ad0-194">Quand l’utilisateur modifie ou supprime une cellule qui fait référence à la fonction.</span><span class="sxs-lookup"><span data-stu-id="60ad0-194">The user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="60ad0-p120">Lorsque l’un des arguments (entrées) de la fonction est modifié. Dans ce cas, un nouvel appel de fonction est déclenché après l’annulation.</span><span class="sxs-lookup"><span data-stu-id="60ad0-p120">When one of the arguments (inputs) for the function changes. In this case, a new function call is triggered following the cancellation.</span></span>

- <span data-ttu-id="60ad0-p121">Lorsque l’utilisateur déclenche le recalcul manuellement. Dans ce cas, un nouvel appel de fonction est déclenché après l’annulation.</span><span class="sxs-lookup"><span data-stu-id="60ad0-p121">When the user triggers recalculation manually. In this case, a new function call is triggered following the cancellation.</span></span>

<span data-ttu-id="60ad0-p122">Pour activer la possibilité d’annuler une fonction, vous devez implémenter un gestionnaire d’annulation dans la fonction JavaScript et spécifier la propriété `"cancelable": true` dans l'objet `options` des métadonnées JSON qui décrit la fonction. Les exemples de code dans la section précédente de cet article fournissent un exemple de ces techniques.</span><span class="sxs-lookup"><span data-stu-id="60ad0-p122">To enable the ability to cancel a function, you must implement a cancellation handler within the JavaScript function and specify the property `"cancelable": true` within the `options` object in the JSON metadata that describes the function. The code samples in the previous section of this article provide an example of these techniques.</span></span>

## <a name="saving-and-sharing-state"></a><span data-ttu-id="60ad0-201">Enregistrement et partage de l'état</span><span class="sxs-lookup"><span data-stu-id="60ad0-201">Saving and sharing state</span></span>

<span data-ttu-id="60ad0-p123">Fonctions personnalisées peuvent enregistrer les données dans les variables globales JavaScript, qui peuvent être utilisés dans les appels suivants. État enregistré est utile lorsque les utilisateurs appellent la même fonction personnalisée à partir de plusieurs cellules, car toutes les instances de la fonction peuvent accéder à l’état. Par exemple, vous pouvez enregistrer les données renvoyées par un appel à une ressource web pour éviter d’effectuer des appels à la même ressource web supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="60ad0-p123">Custom functions can save data in global JavaScript variables. In subsequent calls, your custom function may use the values saved in these variables. Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state. For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="60ad0-p124">L’exemple de code suivant illustre l'implémentation d’une fonction de diffusion en continu de température qui enregistre l’état de manière globale. Notez ce qui suit concernant ce code :</span><span class="sxs-lookup"><span data-stu-id="60ad0-p124">The following code sample shows an implementation of a temperature-streaming function that saves state globally. Note the following about this code:</span></span>

- <span data-ttu-id="60ad0-207">Le `streamTemperature` fonction met à jour la valeur de température qui s’affiche dans la cellule par seconde et qu’il utilise le `savedTemperatures` variable comme source de données.</span><span class="sxs-lookup"><span data-stu-id="60ad0-207">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="60ad0-208">Étant donné que `streamTemperature` est une fonction de diffusion en continu, il implémente un gestionnaire d’annulation qui s’exécute lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="60ad0-208">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="60ad0-209">Si un utilisateur appelle le `streamTemperature` fonction à partir de plusieurs cellules dans Excel, les `streamTemperature` fonction lit les données de la même `savedTemperatures` variable chaque fois qu’elle s’exécute.</span><span class="sxs-lookup"><span data-stu-id="60ad0-209">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="60ad0-210">Le `refreshTemperature` fonction lit la température d’un enregistreur particulier par seconde et stocke le résultat dans le `savedTemperatures` variable.</span><span class="sxs-lookup"><span data-stu-id="60ad0-210">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="60ad0-211">Étant donné que la `refreshTemperature` fonction n’est pas exposée aux utilisateurs finaux dans Excel, il n’a pas besoin d’être enregistré dans le fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="60ad0-211">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

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

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="60ad0-212">Utilisation des plages de données</span><span class="sxs-lookup"><span data-stu-id="60ad0-212">Working with ranges of data</span></span>

<span data-ttu-id="60ad0-p126">Votre fonction personnalisée peut accepter une plage de données comme paramètre d’entrée, ou elle peut renvoyer une plage de données. En JavaScript, une plage de données est représentée sous forme de tableau 2D.</span><span class="sxs-lookup"><span data-stu-id="60ad0-p126">Your custom function may accept a range of data as an input parameter, or it may return a range of data. In JavaScript, a range of data is represented as a 2-dimensional array.</span></span>

<span data-ttu-id="60ad0-p127">Par exemple, supposons que votre fonction renvoie la deuxième valeur la plus élevée prise dans une plage de nombres stockés dans Excel. La fonction suivante accepte le paramètre `values`, qui est de type `Excel.CustomFunctionDimensionality.matrix`. Notez que dans les métadonnées JSON pour cette fonction, vous définissez la propriété `type` du paramètre sur `matrix`.</span><span class="sxs-lookup"><span data-stu-id="60ad0-p127">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel. The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`. Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="handling-errors"></a><span data-ttu-id="60ad0-218">Gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="60ad0-218">Handling errors</span></span>

<span data-ttu-id="60ad0-p128">Lorsque vous créez un complément qui définit des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution. La gestion des erreurs pour les fonctions personnalisées est identique à la [gestion des erreurs pour l’API JavaScript d’Excel dans son ensemble](excel-add-ins-error-handling.md). Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.</span><span class="sxs-lookup"><span data-stu-id="60ad0-p128">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors. Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md). In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="known-issues"></a><span data-ttu-id="60ad0-222">Problèmes connus</span><span class="sxs-lookup"><span data-stu-id="60ad0-222">Known issues</span></span>

- <span data-ttu-id="60ad0-223">Les descriptions de paramètre et les URL d’aide ne sont pas encore utilisées par Excel.</span><span class="sxs-lookup"><span data-stu-id="60ad0-223">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="60ad0-224">Les fonctions personnalisées ne sont actuellement pas disponibles sur Excel pour les clients mobiles.</span><span class="sxs-lookup"><span data-stu-id="60ad0-224">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="60ad0-225">Les fonctions volatiles (celles qui recalculent automatiquement lorsque des modifications de données indépendantes sont effectuées dans la feuille de calcul) ne sont pas encore prises en charge.</span><span class="sxs-lookup"><span data-stu-id="60ad0-225">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="60ad0-226">Le déploiement via le portail d'administration Office 365 et AppSource n'est pas encore activé.</span><span class="sxs-lookup"><span data-stu-id="60ad0-226">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="60ad0-p129">Les fonctions personnalisées dans Excel Online peuvent cesser de fonctionner pendant une session après une période d'inactivité. Actualisez la page du navigateur (F5) et entrez à nouveau une fonction personnalisée pour restaurer la fonction.</span><span class="sxs-lookup"><span data-stu-id="60ad0-p129">Custom functions in Excel Online may stop working during a session after a period of inactivity. Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="60ad0-p130">Il est possible d’avoir le résultat temporaire **#GETTING_DATA** dans la ou les cellules d’une feuille de calcul si vous avez plusieurs compléments s’exécutant dans Excel pour Windows. Fermez toutes les fenêtres d'Excel et redémarrez Excel.</span><span class="sxs-lookup"><span data-stu-id="60ad0-p130">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows. Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="60ad0-p131">Des outils de débogage spécifiques pour les fonctions personnalisées pourraient devenir disponibles à l’avenir. En attendant, vous pouvez déboguer sur Excel Online à l’aide des outils de développement F12. Voir plus de détails dans la section [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="60ad0-p131">Debugging tools specifically for custom functions may be available in the future. In the meantime, you can debug on Excel Online using F12 developer tools. See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>

## <a name="changelog"></a><span data-ttu-id="60ad0-234">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="60ad0-234">Changelog</span></span>

- <span data-ttu-id="60ad0-235">**7 novembre 2017 :** mise à disposition\* de la préversion des fonctions personnalisées et d'exemples</span><span class="sxs-lookup"><span data-stu-id="60ad0-235">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="60ad0-236">**20 novembre 2017**: correction du bogue de compatibilité pour les utilisateurs de la version 8801 et ultérieure</span><span class="sxs-lookup"><span data-stu-id="60ad0-236">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="60ad0-237">**28 novembre 2017 :** mise à disposition\* de la prise en charge de l’annulation sur des fonctions asynchrones (nécessite la modification des fonctions de flux)</span><span class="sxs-lookup"><span data-stu-id="60ad0-237">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="60ad0-238">**7 mai 2018** : support fourni\*pour Mac, Excel Online et les fonctions synchrones en cours de traitement</span><span class="sxs-lookup"><span data-stu-id="60ad0-238">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="60ad0-p132">**20 septembre 2018** : Support fourni pour les fonctions personnalisées à l'exécution de JavaScript. Pour plus d’informations, voir la section [Exécution des fonctions personnalisées d’Excel](custom-functions-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="60ad0-p132">**September 20, 2018**: Shipped support for custom functions JavaScript runtime. For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>

<span data-ttu-id="60ad0-241">\* vers le canal Office Insiders</span><span class="sxs-lookup"><span data-stu-id="60ad0-241">\* to the Office Insiders Channel</span></span>

## <a name="see-also"></a><span data-ttu-id="60ad0-242">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="60ad0-242">See also</span></span>

* [<span data-ttu-id="60ad0-243">Métadonnées des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="60ad0-243">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="60ad0-244">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="60ad0-244">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="60ad0-245">Meilleures pratiques pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="60ad0-245">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="60ad0-246">Didacticiel sur les fonctions personnalisées d’Excel</span><span class="sxs-lookup"><span data-stu-id="60ad0-246">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)