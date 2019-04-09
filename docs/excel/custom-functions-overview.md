---
ms.date: 03/29/2019
description: Créer des fonctions personnalisées dans Excel à l’aide de JavaScript.
title: Créer des fonctions personnalisées dans Excel (aperçu)
localization_priority: Priority
ms.openlocfilehash: 59620b19cb8613e411abb84ed6766da94cae02c4
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/04/2019
ms.locfileid: "31477557"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="4c256-103">Créer des fonctions personnalisées dans Excel (aperçu)</span><span class="sxs-lookup"><span data-stu-id="4c256-103">Create custom functions in Excel (preview)</span></span>

<span data-ttu-id="4c256-104">Les fonctions personnalisées permettent aux développeurs d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément.</span><span class="sxs-lookup"><span data-stu-id="4c256-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="4c256-105">Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="4c256-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="4c256-106">Cet article explique comment créer des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="4c256-106">This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="4c256-107">L’illustration suivante montre un utilisateur final insérant une fonction personnalisée dans une cellule de feuille de calcul Excel.</span><span class="sxs-lookup"><span data-stu-id="4c256-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="4c256-108">Le `CONTOSO.ADD42` fonction personnalisée est conçue pour ajouter 42 à la paire de nombres que spécifie l’utilisateur en tant que paramètres d’entrée de la fonction.</span><span class="sxs-lookup"><span data-stu-id="4c256-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="4c256-109">Le code suivant définit la `ADD42` fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="4c256-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="4c256-110">La section [problèmes connus](#known-issues)plus loin dans cet article indique les limitations en cours de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="4c256-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="4c256-111">Composants d’un projet de complément fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="4c256-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="4c256-112">Si vous utilisez le [générateur de Yo Office](https://github.com/OfficeDev/generator-office) pour créer un projet de complément de fonctions personnalisées Excel, vous constaterez qu’il crée des fichiers qui contrôlent totalement vos fonctions, votre volet des tâches et votre complément.</span><span class="sxs-lookup"><span data-stu-id="4c256-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll find that it creates files which control your functions, your task pane, and your add-in overall.</span></span> <span data-ttu-id="4c256-113">Nous allons vous concentrer sur les fichiers importants pour les fonctions personnalisées :</span><span class="sxs-lookup"><span data-stu-id="4c256-113">We'll concentrate on the files that are important to custom functions:</span></span> 

| <span data-ttu-id="4c256-114">File</span><span class="sxs-lookup"><span data-stu-id="4c256-114">File</span></span> | <span data-ttu-id="4c256-115">Format de fichier</span><span class="sxs-lookup"><span data-stu-id="4c256-115">File format</span></span> | <span data-ttu-id="4c256-116">Description</span><span class="sxs-lookup"><span data-stu-id="4c256-116">Description</span></span> |
|------|-------------|-------------|
| **<span data-ttu-id="4c256-117">./src/functions/functions.js</span><span class="sxs-lookup"><span data-stu-id="4c256-117">./src/functions/functions.js</span></span>**<br/><span data-ttu-id="4c256-118">ou</span><span class="sxs-lookup"><span data-stu-id="4c256-118">or</span></span><br/>**<span data-ttu-id="4c256-119">./src/functions/functions.ts</span><span class="sxs-lookup"><span data-stu-id="4c256-119">./src/functions/functions.ts</span></span>** | <span data-ttu-id="4c256-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="4c256-120">JavaScript</span></span><br/><span data-ttu-id="4c256-121">ou</span><span class="sxs-lookup"><span data-stu-id="4c256-121">or</span></span><br/><span data-ttu-id="4c256-122">TypeScript</span><span class="sxs-lookup"><span data-stu-id="4c256-122">TypeScript</span></span> | <span data-ttu-id="4c256-123">Contient le code qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="4c256-123">Contains the code that defines custom functions.</span></span> |
| **<span data-ttu-id="4c256-124">./src/functions/functions.html</span><span class="sxs-lookup"><span data-stu-id="4c256-124">./src/functions/functions.html</span></span>** | <span data-ttu-id="4c256-125">HTML</span><span class="sxs-lookup"><span data-stu-id="4c256-125">HTML</span></span> | <span data-ttu-id="4c256-126">Fournit une référence&lt;script&gt; au fichier JavaScript qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="4c256-126">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| **<span data-ttu-id="4c256-127">./manifest.xml</span><span class="sxs-lookup"><span data-stu-id="4c256-127">./manifest.xml</span></span>** | <span data-ttu-id="4c256-128">XML</span><span class="sxs-lookup"><span data-stu-id="4c256-128">XML</span></span> | <span data-ttu-id="4c256-129">Spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers JavaScript et HTML qui figurent plus haut dans ce tableau.</span><span class="sxs-lookup"><span data-stu-id="4c256-129">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> <span data-ttu-id="4c256-130">Répertorie également les emplacements des autres fichiers que votre complément pourrait utiliser, tels que les fichiers du volet des tâches et les fichiers de commande.</span><span class="sxs-lookup"><span data-stu-id="4c256-130">It also lists the locations of other files your add-in might make use of, such as the task pane files and command files.</span></span> |

### <a name="script-file"></a><span data-ttu-id="4c256-131">Fichier de script</span><span class="sxs-lookup"><span data-stu-id="4c256-131">Script file</span></span>

<span data-ttu-id="4c256-132">Le fichier de script (**./src/functions/functions.js** ou **./src/functions/functions.ts** dans le projet que crée le générateur de Yo Office) contient le code qui définit des fonctions personnalisées, des commentaires qui définissent la fonction, et associe les noms des fonctions personnalisées à des objets dans le [fichier de métadonnées JSON](#json-metadata-file).</span><span class="sxs-lookup"><span data-stu-id="4c256-132">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span>

<span data-ttu-id="4c256-133">Le code suivant définit la fonction personnalisée `add`, puis spécifie des informations d’association pour la fonction.</span><span class="sxs-lookup"><span data-stu-id="4c256-133">The following code defines the custom function `add`  and then specifies association information for the function.</span></span> <span data-ttu-id="4c256-134">Pour plus d’informations sur l’association de fonctions, voir [Meilleures pratiques des fonctions personnalisées](custom-functions-best-practices.md#associating-function-names-with-json-metadata).</span><span class="sxs-lookup"><span data-stu-id="4c256-134">For more information, see [Custom functions best practices (preview)](custom-functions-best-practices.md#associating-function-names-with-json-metadata).</span></span>

<span data-ttu-id="4c256-135">Le code suivant fournit également des commentaires de code qui définissent la fonction.</span><span class="sxs-lookup"><span data-stu-id="4c256-135">The following code also provides code comments which define the function.</span></span> <span data-ttu-id="4c256-136">Le commentaire obligatoire `@customfunction` est déclaré en premier, pour indiquer qu’il s’agit d’une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="4c256-136">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="4c256-137">Vous pouvez également constater que deux paramètres sont déclarés, `first` et `second`, qui sont suivis de leurs propriétés `description`.</span><span class="sxs-lookup"><span data-stu-id="4c256-137">Additionally, you'll notice two parameters are declared, `first` and `second`, which are followed by their `description` properties.</span></span> <span data-ttu-id="4c256-138">Enfin, une description `returns` est fournie.</span><span class="sxs-lookup"><span data-stu-id="4c256-138">Finally, a `returns` description is given.</span></span> <span data-ttu-id="4c256-139">Pour plus d’informations sur les commentaires requis pour votre fonction personnalisée, voir [Générer des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="4c256-139">For more information about what comments are required for your custom function, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
```

### <a name="manifest-file"></a><span data-ttu-id="4c256-140">Fichier manifeste</span><span class="sxs-lookup"><span data-stu-id="4c256-140">Manifest file</span></span>

<span data-ttu-id="4c256-141">Le fichier manifeste XML pour un complément qui définit les fonctions personnalisées (**./manifest.xml** du projet créé par le Générateur de Yo Office) spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers HTML, JavaScript et JSON.</span><span class="sxs-lookup"><span data-stu-id="4c256-141">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> 

<span data-ttu-id="4c256-142">Le marquage XML suivant présente un exemple des éléments`<ExtensionPoint>` et `<Resources>` que vous devez inclure dans le manifeste d’un complément pour activer les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="4c256-142">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span> <span data-ttu-id="4c256-143">Si vous utilisez le générateur de Yo Office, vos fichiers de fonction personnalisée générés contiennent un fichier manifeste plus complexe que vous pouvez comparer sur [ce dépôt Github](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml).</span><span class="sxs-lookup"><span data-stu-id="4c256-143">If using the Yo Office generator, your generated custom function files will contain a more complex manifest file, which you can compare on t[his Github repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml).</span></span>

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
> <span data-ttu-id="4c256-144">Les fonctions dans Excel sont précédées par l’espace de noms spécifié dans votre fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="4c256-144">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="4c256-145">L’espace de noms d’une fonction vient avant le nom de fonction et les deux sont séparés par un point.</span><span class="sxs-lookup"><span data-stu-id="4c256-145">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="4c256-146">Par exemple, pour appeler la fonction `ADD42` dans la cellule de feuille de calcul Excel, vous saisiriez `=CONTOSO.ADD42`, car `CONTOSO` est l’espace de noms et `ADD42` est le nom de la fonction spécifié dans le fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="4c256-146">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="4c256-147">L’espace de noms est destiné à être utilisé comme identificateur de votre entreprise ou du complément.</span><span class="sxs-lookup"><span data-stu-id="4c256-147">The namespace is intended to be used as an identifier for your company or the add-in.</span></span> <span data-ttu-id="4c256-148">Un espace de noms ne peut contenir que des points et des caractères alphanumériques.</span><span class="sxs-lookup"><span data-stu-id="4c256-148">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="declaring-a-volatile-function"></a><span data-ttu-id="4c256-149">Déclaration d’une fonction volatile</span><span class="sxs-lookup"><span data-stu-id="4c256-149">Declaring a volatile function</span></span>

<span data-ttu-id="4c256-150">Les [fonctions volatiles](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) sont des fonctions dont la valeur change d’un moment à l’autre, même si aucun des arguments de la fonction n’a été modifié.</span><span class="sxs-lookup"><span data-stu-id="4c256-150">[Volatile functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) are functions in which the value changes from moment to moment, even if none of the function's arguments have changed.</span></span> <span data-ttu-id="4c256-151">Ces fonctions sont recalculées à chaque recalcul d’Excel.</span><span class="sxs-lookup"><span data-stu-id="4c256-151">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="4c256-152">Par exemple, imaginons une cellule qui appelle la fonction `NOW`.</span><span class="sxs-lookup"><span data-stu-id="4c256-152">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="4c256-153">Chaque fois que la fonction `NOW` est appelée, elle renvoie automatiquement la date et l’heure actuelles.</span><span class="sxs-lookup"><span data-stu-id="4c256-153">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

<span data-ttu-id="4c256-154">Excel contient plusieurs fonctions volatiles intégrées, comme `RAND` et `TODAY`.</span><span class="sxs-lookup"><span data-stu-id="4c256-154">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="4c256-155">Pour obtenir la liste complète des fonctions volatiles d’Excel, reportez-vous à [Fonctions volatiles et non volatiles](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span><span class="sxs-lookup"><span data-stu-id="4c256-155">For a comprehensive list of Excel's volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="4c256-156">Les fonctions personnalisées permettent de créer vos propres fonctions volatiles, qui peuvent être utiles lors de la gestion des dates, des heures, des nombres aléatoires et de la modélisation.</span><span class="sxs-lookup"><span data-stu-id="4c256-156">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modelling.</span></span> <span data-ttu-id="4c256-157">Par exemple, les simulations Monte Carlo exigent la génération d’entrées aléatoires afin de déterminer une solution optimale.</span><span class="sxs-lookup"><span data-stu-id="4c256-157">For example, Monte Carlo simulations require generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="4c256-158">Pour déclarer une fonction volatile, ajoutez `"volatile": true` au sein de l’objet `options` pour la fonction dans le fichier de métadonnées JSON, comme indiqué dans l’exemple de code suivant.</span><span class="sxs-lookup"><span data-stu-id="4c256-158">To declare a function volatile, add `"volatile": true` within the `options` object  for the function in the JSON metadata file, as shown in the following code sample.</span></span> <span data-ttu-id="4c256-159">Notez qu’une fonction ne peut pas être marquée à la fois `"streaming": true` et `"volatile": true`. Dans le cas où les deux sont marquées comme `true`, l’option volatile est ignorée.</span><span class="sxs-lookup"><span data-stu-id="4c256-159">Note that a function cannot be marked both `"streaming": true` and `"volatile": true`; in the case where both are marked `true` the volatile option will be ignored.</span></span>

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

## <a name="saving-and-sharing-state"></a><span data-ttu-id="4c256-160">Enregistrement et partage d’état</span><span class="sxs-lookup"><span data-stu-id="4c256-160">Saving and sharing state</span></span>

<span data-ttu-id="4c256-161">Les fonctions personnalisées peuvent enregistrer des données dans des variables JavaScript globales, qui peuvent être utilisées dans les appels suivants.</span><span class="sxs-lookup"><span data-stu-id="4c256-161">Custom functions can save data in global JavaScript variables, which can be used in subsequent calls.</span></span> <span data-ttu-id="4c256-162">Un état enregistré est utile lorsque les utilisateurs appellent la même fonction personnalisée à partir de plusieurs cellules, car toutes les instances de la fonction pouvant accéder à l’état.</span><span class="sxs-lookup"><span data-stu-id="4c256-162">Saved state is useful when users call the same custom function from more than one cell, because all instances of the function can access the state.</span></span> <span data-ttu-id="4c256-163">Par exemple, vous pouvez enregistrer les données renvoyées par un appel à une ressource web pour éviter d’effectuer des appels supplémentaires à la même ressource web.</span><span class="sxs-lookup"><span data-stu-id="4c256-163">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="4c256-164">L’exemple de code suivant montre une implémentation d’une fonction de diffusion en continu de la température qui enregistre l’état global.</span><span class="sxs-lookup"><span data-stu-id="4c256-164">The following code sample shows an implementation of a temperature-streaming function that saves state globally.</span></span> <span data-ttu-id="4c256-165">Tenez compte des informations suivantes à propos de ce code :</span><span class="sxs-lookup"><span data-stu-id="4c256-165">Note the following about this code:</span></span>

- <span data-ttu-id="4c256-166">La fonction`streamTemperature`met à jour la valeur de température qui s’affiche dans la cellule chaque seconde et elle utilise la `savedTemperatures` variable en tant que source de données.</span><span class="sxs-lookup"><span data-stu-id="4c256-166">The `streamTemperature` function updates the temperature value that's displayed in the cell every second and it uses the `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="4c256-167">Étant donné que `streamTemperature` est une fonction de diffusion en continu, elle implémente un gestionnaire d’annulation qui s’exécute lorsque la fonction est annulée.</span><span class="sxs-lookup"><span data-stu-id="4c256-167">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="4c256-168">Si un utilisateur appelle la `streamTemperature` fonction à partir de plusieurs cellules dans Excel, la`streamTemperature` fonction lit les données dans la même `savedTemperatures` variable à chaque fois qu’elle s’exécute.</span><span class="sxs-lookup"><span data-stu-id="4c256-168">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="4c256-169">La `refreshTemperature` fonction lit la température d’un thermomètre spécifique à chaque seconde qui passe et stocke le résultat dans la`savedTemperatures`variable.</span><span class="sxs-lookup"><span data-stu-id="4c256-169">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="4c256-170">Étant donné que la `refreshTemperature` fonction n’est pas exposée aux utilisateurs finaux dans Excel, elle n’a pas besoin d’être enregistrée dans le fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="4c256-170">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

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

## <a name="coauthoring"></a><span data-ttu-id="4c256-171">Co-création</span><span class="sxs-lookup"><span data-stu-id="4c256-171">Coauthoring</span></span>

<span data-ttu-id="4c256-172">Excel Online et Excel pour Windows avec un abonnement Office 365 vous permettent de co-créer des documents et cette fonctionnalité est disponible avec les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="4c256-172">Excel Online and Excel for Windows with an Office 365 subscription allow you to coauthor documents and this feature works with custom functions.</span></span> <span data-ttu-id="4c256-173">Si votre classeur utilise une fonction personnalisée, votre collègue sera invité à charger le complément de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="4c256-173">If your workbook uses a custom function, your colleague will be prompted to load the custom function's add-in.</span></span> <span data-ttu-id="4c256-174">Quand vous avez tous les deux chargé le complément, la fonction personnalisée peut partager les résultats via la co-création.</span><span class="sxs-lookup"><span data-stu-id="4c256-174">Once you both have loaded the add-in, the custom function will share results through coauthoring.</span></span>

<span data-ttu-id="4c256-175">Pour plus d’informations sur la co-création, voir [À propos de la co-création dans Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span><span class="sxs-lookup"><span data-stu-id="4c256-175">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="4c256-176">Utilisation des plages de données</span><span class="sxs-lookup"><span data-stu-id="4c256-176">Working with ranges of data</span></span>

<span data-ttu-id="4c256-177">Votre fonction personnalisée peut accepter une plage de données sous la forme d’un paramètre d’entrée, ou il peut renvoyer une plage de données.</span><span class="sxs-lookup"><span data-stu-id="4c256-177">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="4c256-178">Dans JavaScript, une plage de données est représentée sous la forme d’une matrice à deux dimensions.</span><span class="sxs-lookup"><span data-stu-id="4c256-178">In JavaScript, a range of data is represented as a two-dimensional array.</span></span>

<span data-ttu-id="4c256-179">Par exemple, supposons que votre fonction renvoie la seconde valeur la plus élevée à partir d’une plage de nombres stockés dans Excel.</span><span class="sxs-lookup"><span data-stu-id="4c256-179">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="4c256-180">La fonction suivante prend le paramètre `values`, c’est-à-dire un type de `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="4c256-180">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="4c256-181">Notez que dans les métadonnées JSON pour cette fonction, vous devez définir la propriété `type` de paramètre sur `matrix`.</span><span class="sxs-lookup"><span data-stu-id="4c256-181">Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="determine-which-cell-invoked-your-custom-function"></a><span data-ttu-id="4c256-182">Déterminer quelle cellule a appelé votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="4c256-182">Determine which cell invoked your custom function</span></span>

<span data-ttu-id="4c256-183">Dans certains cas, vous devez récupérer l’adresse de la cellule qui a appelé votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="4c256-183">In some cases you'll need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="4c256-184">Cela peut être utile dans les types de scénarios suivants:</span><span class="sxs-lookup"><span data-stu-id="4c256-184">This may be useful in the following types of scenarios:</span></span>

- <span data-ttu-id="4c256-185">Mise en forme de plages: utilisez comme clé la cellule pour stocker des informations dans[AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span><span class="sxs-lookup"><span data-stu-id="4c256-185">Formatting ranges: Use the cell's address as the key to store information in [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="4c256-186">Utilisez ensuite [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) dans Excel pour charger la clé à partir de l’élément `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="4c256-186">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `AsyncStorage`.</span></span>
- <span data-ttu-id="4c256-187">Affichage de valeurs mises en cache : si votre fonction est utilisée en mode hors connexion, affichez les valeurs mises en cache à partir de l’élément `AsyncStorage` à l’aide de `onCalculated`.</span><span class="sxs-lookup"><span data-stu-id="4c256-187">Displaying cached values: If your function is used offline, display stored cached values from `AsyncStorage` using `onCalculated`.</span></span>
- <span data-ttu-id="4c256-188">Rapprochement : utilisez l’adresse de la cellule pour découvrir la cellule d’origine afin de vous aider à réaliser un rapprochement lors du traitement.</span><span class="sxs-lookup"><span data-stu-id="4c256-188">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="4c256-189">Les informations relatives à l’adresse d’une cellule sont exposées uniquement si `requiresAddress` est marqué comme `true` dans le fichier de métadonnées JSON de la fonction.</span><span class="sxs-lookup"><span data-stu-id="4c256-189">The information about a cell's address is exposed only if `requiresAddress` is marked as `true` in the function's JSON metadata file.</span></span> <span data-ttu-id="4c256-190">L’exemple de code suivant illustre ce concept :</span><span class="sxs-lookup"><span data-stu-id="4c256-190">The following sample gives an example of this:</span></span>

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

<span data-ttu-id="4c256-191">Dans le fichier de script (**./src/functions/functions.js** ou **./src/functions/functions.ts**), vous devrez également ajouter une fonction `getAddress` pour trouver l’adresse d’une cellule.</span><span class="sxs-lookup"><span data-stu-id="4c256-191">In the script file (**./src/customfunctions.js** or **./src/customfunctions.ts**), you'll also need to add a `getAddress` function to find a cell's address.</span></span> <span data-ttu-id="4c256-192">Cette fonction peut utiliser des paramètres, comme illustré dans l’exemple suivant en tant que `parameter1`.</span><span class="sxs-lookup"><span data-stu-id="4c256-192">This function may take parameters, as shown in the following sample as `parameter1`.</span></span> <span data-ttu-id="4c256-193">Le dernier paramètre sera toujours `invocationContext`, un objet contenant l’emplacement de la cellule qu’Excel transmet lorsque `requiresAddress` est marqué comme `true` dans votre fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="4c256-193">The last parameter will always be `invocationContext`, an object containing the cell's location that Excel passes down when `requiresAddress` is marked as `true` in your JSON metadata file.</span></span>

```js
function getAddress(parameter1, invocationContext) {
    return invocationContext.address;
}
```

<span data-ttu-id="4c256-194">Par défaut, les valeurs renvoyées par une fonction `getAddress` ont le format suivant : `SheetName!CellNumber`.</span><span class="sxs-lookup"><span data-stu-id="4c256-194">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="4c256-195">Par exemple, si une fonction a été appelée à partir d’une feuille de calcul appelée Dépenses dans la cellule B2, la valeur renvoyée serait `Expenses!B2`.</span><span class="sxs-lookup"><span data-stu-id="4c256-195">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="known-issues"></a><span data-ttu-id="4c256-196">Problèmes connus</span><span class="sxs-lookup"><span data-stu-id="4c256-196">Known issues</span></span>

<span data-ttu-id="4c256-197">Consulter les problèmes connus sur notre[repo GitHub Fonctions Excel Personnalisées](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span><span class="sxs-lookup"><span data-stu-id="4c256-197">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="see-also"></a><span data-ttu-id="4c256-198">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="4c256-198">See also</span></span>

* [<span data-ttu-id="4c256-199">Métadonnées des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="4c256-199">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="4c256-200">Runtime pour les fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="4c256-200">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="4c256-201">Meilleures pratiques des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="4c256-201">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="4c256-202">Fonctions personnalisées changelog</span><span class="sxs-lookup"><span data-stu-id="4c256-202">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="4c256-203">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="4c256-203">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="4c256-204">Débogage des métadonnées des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="4c256-204">Custom functions debugging</span></span>](custom-functions-debugging.md)
