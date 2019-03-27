---
ms.date: 03/19/2019
description: Créer des fonctions personnalisées dans Excel à l’aide de JavaScript.
title: Créer des fonctions personnalisées dans Excel (aperçu)
localization_priority: Priority
ms.openlocfilehash: ac3410267da415c4d567092da2e653fcffd10b72
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870449"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="1dd6d-103">Créer des fonctions personnalisées dans Excel (aperçu)</span><span class="sxs-lookup"><span data-stu-id="1dd6d-103">Create custom functions in Excel (preview)</span></span>

<span data-ttu-id="1dd6d-104">Les fonctions personnalisées permettent aux développeurs d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="1dd6d-105">Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="1dd6d-106">Cet article explique comment créer des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-106">This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="1dd6d-107">L’illustration suivante montre un utilisateur final insérant une fonction personnalisée dans une cellule de feuille de calcul Excel.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="1dd6d-108">Le `CONTOSO.ADD42` fonction personnalisée est conçue pour ajouter 42 à la paire de nombres que spécifie l’utilisateur en tant que paramètres d’entrée de la fonction.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="1dd6d-109">Le code suivant définit la `ADD42` fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="1dd6d-110">La section [problèmes connus](#known-issues)plus loin dans cet article indique les limitations en cours de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="1dd6d-111">Composants d’un projet de complément fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="1dd6d-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="1dd6d-112">Si vous utilisez le [générateur Yo Office](https://github.com/OfficeDev/generator-office) pour créer un projet complément de fonctions personnalisées Excel, vous verrez les fichiers suivants dans le projet crée par le générateur :</span><span class="sxs-lookup"><span data-stu-id="1dd6d-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="1dd6d-113">Fichier</span><span class="sxs-lookup"><span data-stu-id="1dd6d-113">File</span></span> | <span data-ttu-id="1dd6d-114">Format de fichier</span><span class="sxs-lookup"><span data-stu-id="1dd6d-114">File format</span></span> | <span data-ttu-id="1dd6d-115">Description</span><span class="sxs-lookup"><span data-stu-id="1dd6d-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="1dd6d-116">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="1dd6d-116">**./src/customfunctions.js**</span></span><br/><span data-ttu-id="1dd6d-117">ou</span><span class="sxs-lookup"><span data-stu-id="1dd6d-117">or</span></span><br/><span data-ttu-id="1dd6d-118">**./src/customfunctions.ts**</span><span class="sxs-lookup"><span data-stu-id="1dd6d-118">**./src/customfunctions.ts**</span></span> | <span data-ttu-id="1dd6d-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="1dd6d-119">JavaScript</span></span><br/><span data-ttu-id="1dd6d-120">ou</span><span class="sxs-lookup"><span data-stu-id="1dd6d-120">or</span></span><br/><span data-ttu-id="1dd6d-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="1dd6d-121">TypeScript</span></span> | <span data-ttu-id="1dd6d-122">Contient le code qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="1dd6d-123">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="1dd6d-123">**./config/customfunctions.json**</span></span> | <span data-ttu-id="1dd6d-124">JSON</span><span class="sxs-lookup"><span data-stu-id="1dd6d-124">JSON</span></span> | <span data-ttu-id="1dd6d-125">Contient les métadonnées qui décrivent les fonctions personnalisées et permettent à Excel d’enregistrer les fonctions personnalisées et les rendre accessibles aux utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-125">Contains metadata that describes custom functions and enables Excel to register the custom functions and make them available to end users.</span></span> |
| <span data-ttu-id="1dd6d-126">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="1dd6d-126">**./index.html**</span></span> | <span data-ttu-id="1dd6d-127">HTML</span><span class="sxs-lookup"><span data-stu-id="1dd6d-127">HTML</span></span> | <span data-ttu-id="1dd6d-128">Fournit une référence&lt;script&gt; au fichier JavaScript qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="1dd6d-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="1dd6d-129">**./manifest.xml**</span></span> | <span data-ttu-id="1dd6d-130">XML</span><span class="sxs-lookup"><span data-stu-id="1dd6d-130">XML</span></span> | <span data-ttu-id="1dd6d-131">Spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers HTML, JavaScript et JSON qui figurent précédemment dans ce tableau.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="1dd6d-132">Les sections suivantes vous apportent plus d'informations sur ces fichiers.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-132">The following sections provide more information about these files.</span></span>

### <a name="script-file"></a><span data-ttu-id="1dd6d-133">Fichier de script</span><span class="sxs-lookup"><span data-stu-id="1dd6d-133">Script file</span></span>

<span data-ttu-id="1dd6d-134">Le fichier de script (**./src/customfunctions.js** ou **./src/customfunctions.ts** du projet créé par le Générateur de Yo Office) contient le code qui définit les fonctions personnalisées et mappe les noms des fonctions personnalisées aux objets dans le [fichier de métadonnées JSON](#json-metadata-file).</span><span class="sxs-lookup"><span data-stu-id="1dd6d-134">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="1dd6d-135">Par exemple, le code suivant définit les fonctions personnalisées `add` et `increment`indique ensuite les informations de mappage pour les deux fonctions.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-135">For example, the following code defines the custom functions `add` and `increment` and then specifies association information for both functions.</span></span> <span data-ttu-id="1dd6d-136">La fonction`add` mappée à l’objet dans le fichier de métadonnées JSON où la valeur de la `id` propriété est **AJOUTER**et la fonction`increment`mappée à l’objet dans le fichier de métadonnées dans laquelle la valeur de la `id` propriété est **INCRÉMENT**.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-136">The `add` function is associated with the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is associated with the object in the metadata file where the value of the `id` property is **INCREMENT**.</span></span> <span data-ttu-id="1dd6d-137">Voir [Recommandations fonctions personnalisées](custom-functions-best-practices.md#associating-function-names-with-json-metadata) pour plus d’informations sur le mappage des noms de fonction dans le fichier de script pour objets dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-137">See [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata) for more information about associating function names in the script file to objects in the JSON metadata file.</span></span>

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

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
 CustomFunctions.associate("INCREMENT", increment);
```

### <a name="json-metadata-file"></a><span data-ttu-id="1dd6d-138">Fichier de métadonnées JSON</span><span class="sxs-lookup"><span data-stu-id="1dd6d-138">JSON metadata file</span></span>

<span data-ttu-id="1dd6d-139">Le fichier de métadonnées fonctions personnalisées (**./config/customfunctions.json** du projet créé par le Générateur de Yo Office) fournit les informations dont Excel a besoin pour enregistrer les fonctions personnalisées et les rendre disponibles aux utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-139">The custom functions metadata file (**./config/customfunctions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users.</span></span> <span data-ttu-id="1dd6d-140">Les fonctions personnalisées sont enregistrées lorsqu’un utilisateur lance un complément pour la première fois.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-140">Custom functions are registered when a user runs an add-in for the first time.</span></span> <span data-ttu-id="1dd6d-141">Après cela, elles sont disponibles pour cet utilisateur depuis tous les classeurs (c'est-à-dire pas seulement dans le classeur dans lequel le complément est initialement exécuté.)</span><span class="sxs-lookup"><span data-stu-id="1dd6d-141">After that, they are available to that same user in all workbooks (i.e., not only in the workbook where the add-in initially ran.)</span></span>

> [!TIP]
> <span data-ttu-id="1dd6d-142">Les paramètres du serveur qui héberge le fichier JSON doivent avoir [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) activée afin que les fonctions personnalisées s’exécutent correctement dans Excel Online.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-142">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="1dd6d-143">Le code suivant de **customfunctions.json** spécifie les métadonnées pour la `add` fonction et la `increment` fonction qui ont été décrites précédemment.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-143">The following code in **customfunctions.json** specifies the metadata for the `add` function and the `increment` function that were described previously.</span></span> <span data-ttu-id="1dd6d-144">Le tableau qui suit cet exemple de code fournit des informations détaillées sur les propriétés individuelles au sein de cet objet JSON.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-144">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span> <span data-ttu-id="1dd6d-145">Voir [Recommandations fonctions personnalisées](custom-functions-best-practices.md#associating-function-names-with-json-metadata) pour plus d’informations sur la spécification de la valeur de `id` et les propriétés`name`dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-145">See [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

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

<span data-ttu-id="1dd6d-146">Le tableau suivant répertorie les propriétés généralement présentes dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-146">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="1dd6d-147">Pour plus d’informations sur le fichier de métadonnées JSON, voir [métadonnées fonctions personnalisées](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="1dd6d-147">For more detailed information about the JSON metadata file, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="1dd6d-148">Propriété</span><span class="sxs-lookup"><span data-stu-id="1dd6d-148">Property</span></span>  | <span data-ttu-id="1dd6d-149">Description</span><span class="sxs-lookup"><span data-stu-id="1dd6d-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="1dd6d-150">Un ID unique pour la fonction.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-150">A unique ID for the function.</span></span> <span data-ttu-id="1dd6d-151">Cet ID peut contenir uniquement des points et caractères alphanumériques et ne doit pas être modifié une fois défini.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-151">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="1dd6d-152">Nom de la fonction que voit l’utilisateur final dans Excel.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-152">Name of the function that the end user sees in Excel.</span></span> <span data-ttu-id="1dd6d-153">Dans Excel, ce nom de la fonction sera précédé de l’espace de noms de fonctions personnalisées spécifié dans le [fichier manifeste XML](#manifest-file).</span><span class="sxs-lookup"><span data-stu-id="1dd6d-153">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the [XML manifest file](#manifest-file).</span></span> |
| `helpUrl` | <span data-ttu-id="1dd6d-154">URL de la page qui s’affiche quand un utilisateur demande de l’aide.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-154">URL for the page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="1dd6d-155">Descriptif de la fonction.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-155">Describes what the function does.</span></span> <span data-ttu-id="1dd6d-156">Cette valeur apparaît comme une info-bulle lorsque la fonction est l’élément sélectionné dans le menu de saisie semi-automatique des formules dans Excel.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-156">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="1dd6d-157">Objet qui définit le type d’informations renvoyées par la fonction.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-157">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="1dd6d-158">Pour plus d’informations sur cet objet, voir [résultat](custom-functions-json.md#result).</span><span class="sxs-lookup"><span data-stu-id="1dd6d-158">For detailed information about this object, see [result](custom-functions-json.md#result).</span></span> |
| `parameters` | <span data-ttu-id="1dd6d-159">Tableau qui définit les paramètres d’entrée de la fonction.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-159">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="1dd6d-160">Pour plus d’informations sur cet objet, voir [paramètres](custom-functions-json.md#parameters).</span><span class="sxs-lookup"><span data-stu-id="1dd6d-160">For detailed information about this object, see [parameters](custom-functions-json.md#parameters).</span></span> |
| `options` | <span data-ttu-id="1dd6d-161">Vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-161">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="1dd6d-162">Pour plus d’informations sur l’utilisation de cette propriété, consultez les sections [Fonctions de diffusion en continu](#streaming-functions) et [Annulation d’une fonction](#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="1dd6d-162">For more information about how this property can be used, see [Streaming functions](#streaming-functions) and [canceling a function](#canceling-a-function).</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="1dd6d-163">Fichier manifeste</span><span class="sxs-lookup"><span data-stu-id="1dd6d-163">Manifest file</span></span>

<span data-ttu-id="1dd6d-164">Le fichier manifeste XML pour un complément qui définit les fonctions personnalisées (**./manifest.xml** du projet créé par le Générateur de Yo Office) spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers HTML, JavaScript et JSON.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-164">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="1dd6d-165">Le balisage XML suivant montre un exemple des éléments`<ExtensionPoint>` et `<Resources>` que vous devez inclure dans manifeste d’un complément pour activer les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-165">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span>  

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8081/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="JS-URL"/>
            </Script>
            <Page>
              <SourceLocation resid="HTML-URL"/>
            </Page>
            <Metadata>
              <SourceLocation resid="JSON-URL"/>
            </Metadata>
            <Namespace resid="namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
      </Host>
    </Hosts>
    <Resources>
      <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://localhost:8081/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://localhost:8081/dist/win32/ship/index.win32.bundle"/>
        <bt:Url id="HTML-URL" DefaultValue="https://localhost:8081/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
      </bt:ShortStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="1dd6d-166">Les fonctions dans Excel sont précédées par l’espace de noms spécifié dans votre fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-166">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="1dd6d-167">L’espace de noms d’une fonction vient avant le nom de fonction et les deux sont séparés par un point.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-167">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="1dd6d-168">Par exemple, pour appeler la fonction `ADD42` dans la cellule de feuille de calcul Excel, vous saisiriez `=CONTOSO.ADD42`, car `CONTOSO` est l’espace de noms et `ADD42` est le nom de la fonction spécifié dans le fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-168">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="1dd6d-169">L’espace de noms est destiné à être utilisé comme identificateur de votre entreprise ou du complément.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-169">The namespace is intended to be used as an identifier for your company or the add-in.</span></span> <span data-ttu-id="1dd6d-170">Un espace de noms ne peut contenir que des points et des caractères alphanumériques.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-170">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="declaring-a-volatile-function"></a><span data-ttu-id="1dd6d-171">Déclaration d’une fonction volatile</span><span class="sxs-lookup"><span data-stu-id="1dd6d-171">Declaring a volatile function</span></span>

<span data-ttu-id="1dd6d-172">Les [fonctions volatiles](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) sont des fonctions dont la valeur change d’un moment à l’autre, même si aucun des arguments de la fonction n’a été modifié.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-172">[Volatile functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) are functions in which the value changes from moment to moment, even if none of the function's arguments have changed.</span></span> <span data-ttu-id="1dd6d-173">Ces fonctions sont recalculées à chaque recalcul d’Excel.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-173">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="1dd6d-174">Par exemple, imaginons une cellule qui appelle la fonction `NOW`.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-174">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="1dd6d-175">Chaque fois que la fonction `NOW` est appelée, elle renvoie automatiquement la date et l’heure actuelles.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-175">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

<span data-ttu-id="1dd6d-176">Excel contient plusieurs fonctions volatiles intégrées, comme `RAND` et `TODAY`.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-176">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="1dd6d-177">Pour obtenir la liste complète des fonctions volatiles d’Excel, reportez-vous à [Fonctions volatiles et non volatiles](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span><span class="sxs-lookup"><span data-stu-id="1dd6d-177">For a comprehensive list of Excel's volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="1dd6d-178">Les fonctions personnalisées permettent de créer vos propres fonctions volatiles, qui peuvent être utiles lors de la gestion des dates, des heures, des nombres aléatoires et de la modélisation.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-178">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modelling.</span></span> <span data-ttu-id="1dd6d-179">Par exemple, les simulations Monte Carlo exigent la génération d’entrées aléatoires afin de déterminer une solution optimale.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-179">For example, Monte Carlo simulations require generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="1dd6d-180">Pour déclarer une fonction volatile, ajoutez `"volatile": true` au sein de l’objet `options` pour la fonction dans le fichier de métadonnées JSON, comme indiqué dans l’exemple de code suivant.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-180">To declare a function volatile, add `"volatile": true` within the `options` object  for the function in the JSON metadata file, as shown in the following code sample.</span></span> <span data-ttu-id="1dd6d-181">Notez qu’une fonction ne peut pas être marquée à la fois `"streaming": true` et `"volatile": true`. Dans le cas où les deux sont marquées comme `true`, l’option volatile est ignorée.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-181">Note that a function cannot be marked both `"streaming": true` and `"volatile": true`; in the case where both are marked `true` the volatile option will be ignored.</span></span>

```json
{
 "id": "TOMORROW",
  "name": "TOMORROW",
  "description":  "Returns tomorrow’s date",
  "helpUrl": "http://www.contoso.com",
  "result": {
      "type": "string",
      "dimensionality": "scalar"
  },
  "options": {
      "volatile": true
  }
}
```

## <a name="saving-and-sharing-state"></a><span data-ttu-id="1dd6d-182">Enregistrement et partage d’état</span><span class="sxs-lookup"><span data-stu-id="1dd6d-182">Saving and sharing state</span></span>

<span data-ttu-id="1dd6d-183">Les fonctions personnalisées peuvent enregistrer des données dans des variables JavaScript globales, qui peuvent être utilisées dans les appels suivants.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-183">Custom functions can save data in global JavaScript variables, which can be used in subsequent calls.</span></span> <span data-ttu-id="1dd6d-184">Un état enregistré est utile lorsque les utilisateurs appellent la même fonction personnalisée à partir de plusieurs cellules, car toutes les instances de la fonction pouvant accéder à l’état.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-184">Saved state is useful when users call the same custom function from more than one cell, because all instances of the function can access the state.</span></span> <span data-ttu-id="1dd6d-185">Par exemple, vous pouvez enregistrer les données renvoyées par un appel à une ressource web pour éviter d’effectuer des appels supplémentaires à la même ressource web.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-185">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="1dd6d-186">L’exemple de code suivant montre une implémentation d’une fonction de diffusion en continu de la température qui enregistre l’état global.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-186">The following code sample shows an implementation of a temperature-streaming function that saves state globally.</span></span> <span data-ttu-id="1dd6d-187">Tenez compte des informations suivantes à propos de ce code :</span><span class="sxs-lookup"><span data-stu-id="1dd6d-187">Note the following about this code:</span></span>

- <span data-ttu-id="1dd6d-188">La fonction`streamTemperature`met à jour la valeur de température qui s’affiche dans la cellule chaque seconde et elle utilise la `savedTemperatures` variable en tant que source de données.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-188">The `streamTemperature` function updates the temperature value that's displayed in the cell every second and it uses the `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="1dd6d-189">Étant donné que `streamTemperature` est une fonction de diffusion en continu, elle implémente un gestionnaire d’annulation qui s’exécute lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-189">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="1dd6d-190">Si un utilisateur appelle la `streamTemperature` fonction à partir de plusieurs cellules dans Excel, la`streamTemperature` fonction lit les données dans la même `savedTemperatures` variable à chaque fois qu’elle s’exécute.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-190">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="1dd6d-191">La `refreshTemperature` fonction lit la température d’un thermomètre spécifique à chaque seconde qui passe et stocke le résultat dans la`savedTemperatures`variable.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-191">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="1dd6d-192">Étant donné que la `refreshTemperature` fonction n’est pas exposée aux utilisateurs finaux dans Excel, elle n’a pas besoin d’être enregistrée dans le fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-192">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

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

## <a name="coauthoring"></a><span data-ttu-id="1dd6d-193">Co-création</span><span class="sxs-lookup"><span data-stu-id="1dd6d-193">Coauthoring</span></span>

<span data-ttu-id="1dd6d-194">Excel Online et Excel pour Windows avec un abonnement Office 365 vous permettent de co-créer des documents et cette fonctionnalité est disponible avec les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-194">Excel Online and Excel for Windows with an Office 365 subscription allow you to coauthor documents and this feature works with custom functions.</span></span> <span data-ttu-id="1dd6d-195">Si votre classeur utilise une fonction personnalisée, votre collègue sera invité à charger le complément de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-195">If your workbook uses a custom function, your colleague will be prompted to load the custom function's add-in.</span></span> <span data-ttu-id="1dd6d-196">Quand vous avez tous les deux chargé le complément, la fonction personnalisée peut partager les résultats via la co-création.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-196">Once you both have loaded the add-in, the custom function will share results through coauthoring.</span></span>

<span data-ttu-id="1dd6d-197">Pour plus d’informations sur la co-création, voir [À propos de la co-création dans Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span><span class="sxs-lookup"><span data-stu-id="1dd6d-197">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="1dd6d-198">Utilisation des plages de données</span><span class="sxs-lookup"><span data-stu-id="1dd6d-198">Working with ranges of data</span></span>

<span data-ttu-id="1dd6d-199">Votre fonction personnalisée peut accepter une plage de données sous la forme d’un paramètre d’entrée, ou il peut renvoyer une plage de données.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-199">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="1dd6d-200">Dans JavaScript, une plage de données est représentée sous la forme d’une matrice à deux dimensions.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-200">In JavaScript, a range of data is represented as a two-dimensional array.</span></span>

<span data-ttu-id="1dd6d-201">Par exemple, supposons que votre fonction renvoie la seconde valeur la plus élevée à partir d’une plage de nombres stockés dans Excel.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-201">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="1dd6d-202">La fonction suivante prend le paramètre `values`, c’est-à-dire un type de `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-202">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="1dd6d-203">Notez que dans les métadonnées JSON pour cette fonction, vous devez définir la propriété `type` de paramètre sur `matrix`.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-203">Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="determine-which-cell-invoked-your-custom-function"></a><span data-ttu-id="1dd6d-204">Déterminer quelle cellule a appelé votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-204">Determine which cell invoked your custom function</span></span>

<span data-ttu-id="1dd6d-205">Dans certains cas, vous devez récupérer l’adresse de la cellule qui a appelé votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-205">In some cases you'll need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="1dd6d-206">Cela peut être utile dans les types de scénarios suivants:</span><span class="sxs-lookup"><span data-stu-id="1dd6d-206">This may be useful in the following types of scenarios:</span></span>

- <span data-ttu-id="1dd6d-207">Mise en forme de plages: utilisez comme clé la cellule pour stocker des informations dans[AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span><span class="sxs-lookup"><span data-stu-id="1dd6d-207">Formatting ranges: Use the cell's address as the key to store information in [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="1dd6d-208">Utilisez ensuite [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) dans Excel pour charger la clé à partir de l’élément `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-208">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `AsyncStorage`.</span></span>
- <span data-ttu-id="1dd6d-209">Affichage de valeurs mises en cache : si votre fonction est utilisée en mode hors connexion, affichez les valeurs mises en cache à partir de l’élément `AsyncStorage` à l’aide de `onCalculated`.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-209">Displaying cached values: If your function is used offline, display stored cached values from `AsyncStorage` using `onCalculated`.</span></span>
- <span data-ttu-id="1dd6d-210">Rapprochement : utilisez l’adresse de la cellule pour découvrir la cellule d’origine afin de vous aider à réaliser un rapprochement lors du traitement.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-210">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="1dd6d-211">Les informations relatives à l’adresse d’une cellule sont exposées uniquement si `requiresAddress` est marqué comme `true` dans le fichier de métadonnées JSON de la fonction.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-211">The information about a cell's address is exposed only if `requiresAddress` is marked as `true` in the function's JSON metadata file.</span></span> <span data-ttu-id="1dd6d-212">L’exemple de code suivant illustre ce concept :</span><span class="sxs-lookup"><span data-stu-id="1dd6d-212">The following sample gives an example of this:</span></span>

```JSON
{
   "id": "ADDTIME",
   "name": "ADDTIME",
   "description": "Display current date and add the amount of hours to it designated by the parameter",
   "helpUrl": "http://www.contoso.com",
   "result": {
      "type": "number",
      "dimensionality": "scalar"
   },
   "parameters": [
      {
         "name": "Additional time",
         "description": "Amount of hours to increase current date by",
         "type": "number",
         "dimensionality": "scalar"
      }
   ],
   "options": {
      "requiresAddress": true
   }
}
```

<span data-ttu-id="1dd6d-213">Dans le fichier de script (**./src/customfunctions.js** ou **./src/customfunctions.ts**), vous devrez également ajouter une fonction `getAddress` pour trouver l’adresse d’une cellule.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-213">In the script file (**./src/customfunctions.js** or **./src/customfunctions.ts**), you'll also need to add a `getAddress` function to find a cell's address.</span></span> <span data-ttu-id="1dd6d-214">Cette fonction peut utiliser des paramètres, comme illustré dans l’exemple suivant en tant que `parameter1`.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-214">This function may take parameters, as shown in the following sample as `parameter1`.</span></span> <span data-ttu-id="1dd6d-215">Le dernier paramètre sera toujours `invocationContext`, un objet contenant l’emplacement de la cellule qu’Excel transmet lorsque `requiresAddress` est marqué comme `true` dans votre fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-215">The last parameter will always be `invocationContext`, an object containing the cell's location that Excel passes down when `requiresAddress` is marked as `true` in your JSON metadata file.</span></span>

```js
function getAddress(parameter1, invocationContext) {
    return invocationContext.address;
}
```

<span data-ttu-id="1dd6d-216">Par défaut, les valeurs renvoyées par une fonction `getAddress` ont le format suivant : `SheetName!CellNumber`.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-216">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="1dd6d-217">Par exemple, si une fonction a été appelée à partir d’une feuille de calcul appelée Dépenses dans la cellule B2, la valeur renvoyée serait `Expenses!B2`.</span><span class="sxs-lookup"><span data-stu-id="1dd6d-217">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="known-issues"></a><span data-ttu-id="1dd6d-218">Problèmes connus</span><span class="sxs-lookup"><span data-stu-id="1dd6d-218">Known issues</span></span>

<span data-ttu-id="1dd6d-219">Consulter les problèmes connus sur notre[repo GitHub Fonctions Excel Personnalisées](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span><span class="sxs-lookup"><span data-stu-id="1dd6d-219">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="see-also"></a><span data-ttu-id="1dd6d-220">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="1dd6d-220">See also</span></span>

* [<span data-ttu-id="1dd6d-221">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="1dd6d-221">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="1dd6d-222">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="1dd6d-222">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="1dd6d-223">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="1dd6d-223">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="1dd6d-224">Fonctions personnalisées changelog</span><span class="sxs-lookup"><span data-stu-id="1dd6d-224">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="1dd6d-225">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="1dd6d-225">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
