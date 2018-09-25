---
ms.date: 09/20/2018
description: Créez une fonction personnalisée dans Excel à l’aide de JavaScript.
title: Créer des fonctions personnalisées dans Excel (Préversion)
ms.openlocfilehash: b214329fe50955d0f39d50f674152f475ca24b4d
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/25/2018
ms.locfileid: "25005042"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="5c4b3-103">Créer des fonctions personnalisées dans Excel (Préversion)</span><span class="sxs-lookup"><span data-stu-id="5c4b3-103">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="5c4b3-104">Les fonctions personnalisées permettent aux développeurs d’ajouter de nouvelles fonctions à Excel en définissant ces fonctions en JavaScript dans le cadre d’un complément.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="5c4b3-105">Les utilisateurs Excel peuvent accéder aux fonctions personnalisées comme toute autre fonction native dans Excel (par exemple, `SUM()`).</span><span class="sxs-lookup"><span data-stu-id="5c4b3-105">Users within Excel can access custom functions like any other native function in Excel (such as `SUM()`).</span></span> <span data-ttu-id="5c4b3-106">Cet article décrit comment créer des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-106">This article explains how to create custom functions in Excel.</span></span>

<span data-ttu-id="5c4b3-107">L’illustration suivante montre un utilisateur insérant une fonction personnalisée dans une cellule d’une feuille de calcul Excel.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="5c4b3-108">La fonction personnalisée `CONTOSO.ADD42` est conçue pour ajouter 42 à la paire de nombres spécifiée par l’utilisateur comme paramètres d’entrée de la fonction.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="5c4b3-109">Le code suivant définit la fonction personnalisée `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-109">The following code defines the `ADD42` custom function.</span></span>

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

<span data-ttu-id="5c4b3-110">Les fonctions personnalisées sont désormais disponibles en préversion pour développeur sur Windows, Mac et Excel Online.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-110">Custom functions are now available in Developer Preview on Windows, Mac, and Excel Online.</span></span> <span data-ttu-id="5c4b3-111">Pour les essayer, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="5c4b3-111">To try them, complete these steps:</span></span>

1. <span data-ttu-id="5c4b3-112">Installez Office (version 10827 sur Windows ou 13.329 sur Mac) et participez au programme [Office Insider](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="5c4b3-112">Install Office (build 9325 on Windows or 13.329 on Mac) and join the [Office Insider](https://products.office.com/office-insider) program.</span></span> <span data-ttu-id="5c4b3-113">Vous devez rejoindre le programme Office Insider pour pouvoir accéder aux fonctions personnalisées ; actuellement, les fonctions personnalisées sont désactivées dans toutes les versions d’Office, sauf si vous êtes membre du programme Office Insider.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-113">You must join the Office Insider program in order to have access to custom functions; currently, custom functions are disabled across all Office builds unless you are a member of the Office Insider program.</span></span>

2. <span data-ttu-id="5c4b3-114">Utilisez [Yo Office](https://github.com/OfficeDev/generator-office) pour créer un projet de complément Fonctions Personnalisées Excel, puis suivez les instructions indiquées dans [OfficeDev/Excel-Custom-Functions README](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) pour utiliser le projet.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-114">Use [Yo Office](https://github.com/OfficeDev/generator-office) to create an Excel Custom Functions add-in project, and then follow the instructions in the [OfficeDev/Excel-Custom-Functions README](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) to use the project.</span></span>

3. <span data-ttu-id="5c4b3-115">Saisissez `=CONTOSO.ADD42(1,2)` dans une cellule d’une feuille de calcul Excel et appuyez sur **Entrée** pour exécuter la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-115">Type `=CONTOSO.ADD42(1,2)` into any cell, and press **Enter** to run the custom function.</span></span>

> [!NOTE]
> <span data-ttu-id="5c4b3-116">Plus loin dans cet article, la section [Problèmes connus](#known-issues) indique les limites actuelles des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-116">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="learn-the-basics"></a><span data-ttu-id="5c4b3-117">Notions fondamentales</span><span class="sxs-lookup"><span data-stu-id="5c4b3-117">Learn the basics</span></span>

<span data-ttu-id="5c4b3-118">Dans le projet de fonctions personnalisées que vous avez créé à l’aide de [Yo Office](https://github.com/OfficeDev/generator-office), vous verrez les fichiers suivants :</span><span class="sxs-lookup"><span data-stu-id="5c4b3-118">In the custom functions project that you've created using [Yo Office](https://github.com/OfficeDev/generator-office), you’ll see the following files:</span></span>

| <span data-ttu-id="5c4b3-119">Fichier</span><span class="sxs-lookup"><span data-stu-id="5c4b3-119">File</span></span> | <span data-ttu-id="5c4b3-120">Format de fichier</span><span class="sxs-lookup"><span data-stu-id="5c4b3-120">File format</span></span> | <span data-ttu-id="5c4b3-121">Description</span><span class="sxs-lookup"><span data-stu-id="5c4b3-121">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="5c4b3-122">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="5c4b3-122">**./src/customfunctions.js**</span></span> | <span data-ttu-id="5c4b3-123">JavaScript</span><span class="sxs-lookup"><span data-stu-id="5c4b3-123">JavaScript</span></span> | <span data-ttu-id="5c4b3-124">Contient le code qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-124">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="5c4b3-125">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="5c4b3-125">**./config/customfunctions.json**</span></span> | <span data-ttu-id="5c4b3-126">JSON</span><span class="sxs-lookup"><span data-stu-id="5c4b3-126">JSON</span></span> | <span data-ttu-id="5c4b3-127">Contient des métadonnées qui décrivent les fonctions personnalisées et permettent à Excel d’enregistrer les fonctions personnalisées afin de les rendre disponibles pour les utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-127">Contains metadata that describes custom functions and enables Excel to register the custom functions in order to make them available to end-users.</span></span> |
| <span data-ttu-id="5c4b3-128">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="5c4b3-128">**./index.html**</span></span> | <span data-ttu-id="5c4b3-129">HTML</span><span class="sxs-lookup"><span data-stu-id="5c4b3-129">HTML</span></span> | <span data-ttu-id="5c4b3-130">Fournit une référence de &lt;script&gt; pour le fichier JavaScript qui définit les fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-130">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="5c4b3-131">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="5c4b3-131">**Manifest.xml**</span></span> | <span data-ttu-id="5c4b3-132">XML</span><span class="sxs-lookup"><span data-stu-id="5c4b3-132">XML</span></span> | <span data-ttu-id="5c4b3-133">Spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers JavaScript, JSON et HTML, répertoriés précédemment dans ce tableau.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-133">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

### <a name="manifest-file-manifestxml"></a><span data-ttu-id="5c4b3-134">Fichier manifeste (./manifest.xml)</span><span class="sxs-lookup"><span data-stu-id="5c4b3-134">Manifest file (manifest.xml)</span></span>

<span data-ttu-id="5c4b3-135">Le fichier manifeste XML d’un complément qui définit les fonctions personnalisées spécifie également l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers JavaScript, JSON et HTML.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-135">The XML manifest file for an add-in that defines custom functions specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="5c4b3-136">Le code XML suivant montre un exemple des éléments `<ExtensionPoint>` et `<Resources>` que vous devez inclure dans le manifeste d’un complément pour permettre à Excel d’exécuter des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-136">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest in order to enable Excel to run custom functions.</span></span>  

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
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. For example, a function named "ADD42" is invoked as `=CONTOSO.ADD42` in Excel.-->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="5c4b3-137">Les fonctions dans Excel sont précédées par l’espace de noms spécifié dans votre fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-137">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="5c4b3-138">L’espace de noms d’une fonction précède le nom de la fonction, et ils sont séparés par un point.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-138">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="5c4b3-139">Par exemple, pour appeler la fonction `ADD42()` dans la cellule d’une feuille de calcul Excel, vous devez taper `=CONTOSO.ADD42`, puisque CONTOSO est l’espace de noms et `ADD42` est le nom de la fonction spécifiée dans le fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-139">For example, to call the function `ADD42()` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because CONTOSO is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="5c4b3-140">L’espace de noms est destiné à être utilisé comme identificateur pour votre entreprise ou le complément.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-140">The prefix is intended to be used as an identifier for your add-in.</span></span> 

### <a name="json-file-configcustomfunctionsjson"></a><span data-ttu-id="5c4b3-141">Fichier JSON (./config/customfunctions.json)</span><span class="sxs-lookup"><span data-stu-id="5c4b3-141">JSON file (./config/customfunctions.json)</span></span>

<span data-ttu-id="5c4b3-142">Un fichier de métadonnées des fonctions personnalisées fournit les informations dont Excel a besoin pour inscrire les fonctions personnalisées et les rendre disponibles pour les utilisateurs finaux.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-142">A custom functions metadata file provides the information that Excel requires to register the custom functions and make them available to end-users.</span></span> <span data-ttu-id="5c4b3-143">Les fonctions personnalisées sont enregistrées lorsqu’un utilisateur exécute un complément pour la première fois.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-143">The custom functions are registered when a user runs the add-in for the first time.</span></span> <span data-ttu-id="5c4b3-144">Après cela, elles sont disponibles pour cet utilisateur dans tous les classeurs (autrement dit, pas seulement dans le classeur dans lequel le complément a été exécuté pour la première fois.)</span><span class="sxs-lookup"><span data-stu-id="5c4b3-144">After that, they are available, for that same user, in all workbooks (not only the one where the add-in ran initially.)</span></span>

> [!TIP]
> <span data-ttu-id="5c4b3-145">Parmi les paramètres de serveur sur le serveur qui héberge le fichier JSON, [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) doit être activé pour que les fonctions personnalisées fonctionnent correctement dans Excel Online.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-145">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="5c4b3-146">Le code suivant dans **customfunctions.json** spécifie les métadonnées pour la fonction `ADD42` décrite précédemment dans cet article.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-146">The following code in **customfunctions.json** specifies the metadata for the `ADD42` function that was described previously in this article.</span></span> <span data-ttu-id="5c4b3-147">Ces métadonnées définissent le nom, la description, la valeur renvoyée, les paramètres d’entrée de la fonction, et plus.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-147">This metadata defines the function's name, description, return value, input parameters, and more.</span></span> <span data-ttu-id="5c4b3-148">Le tableau qui suit cet exemple de code fournit des informations détaillées sur les propriétés individuelles dans cet objet JSON.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-148">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span>

```json
{
    "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
    "functions": [
        {
            "id": "ADD42",
            "name": "ADD42",
            "description":  "adds 42 to the input numbers",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [                {
                    "name": "number 1",
                    "description": "the first number to be added",
                    "type": "number",
                    "dimensionality": "scalar"
                },
                {
                    "name": "number 2",
                    "description": "the second number to be added",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ],
        }
    ]
}
```

<span data-ttu-id="5c4b3-149">Le tableau suivant répertorie les propriétés qui sont généralement présentes dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-149">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="5c4b3-150">Pour plus d’informations sur le fichier de métadonnées JSON, y compris sur des options qui n’ont pas été utilisées dans l’exemple précédent, voir [Métadonnées des fonctions personnalisées](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="5c4b3-150">For more detailed information about the JSON metadata file, including options not used in the previous example, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="5c4b3-151">Propriété</span><span class="sxs-lookup"><span data-stu-id="5c4b3-151">Property</span></span>  | <span data-ttu-id="5c4b3-152">Description</span><span class="sxs-lookup"><span data-stu-id="5c4b3-152">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="5c4b3-153">ID unique de la fonction.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-153">A unique ID for the group.</span></span> <span data-ttu-id="5c4b3-154">Cet ID ne doit pas être modifié après sa définition.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-154">This ID should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="5c4b3-155">Nom de la fonction qui est affichée dans le menu de saisie semi-automatique quand un utilisateur tape une formule dans une cellule.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-155">Name of the function that is shown in the autocomplete menu as a user types a formula within a cell.</span></span> <span data-ttu-id="5c4b3-156">Dans le menu de saisie semi-automatique, cette valeur sera préfixée par l’espace de noms des fonctions personnalisées spécifié dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-156">In the autocomplete menu, this value will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `helpUrl` | <span data-ttu-id="5c4b3-157">URL de la page qui s’affiche lorsqu’un utilisateur demande de l’aide.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-157">Url for a page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="5c4b3-158">Décrit ce que fait la fonction.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-158">Describes what the function does.</span></span> <span data-ttu-id="5c4b3-159">Cette valeur s’affiche comme une info-bulle lorsque la fonction est l’élément sélectionné dans le menu de saisie semi-automatique dans Excel.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-159">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="5c4b3-160">Objet qui définit le type de l’information renvoyée par la fonction.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-160">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="5c4b3-161">La valeur de la propriété enfant `type` peut être **string**, **number**ou **boolean**.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-161">The value of the `type` child property can be **string**, **number**, or **boolean**.</span></span> <span data-ttu-id="5c4b3-162">La valeur de la propriété enfant `dimensionality` peut être **scalar** ou **matrix** (tableau à deux dimensions des valeurs du `type` spécifié).</span><span class="sxs-lookup"><span data-stu-id="5c4b3-162">The `dimensionality` property can be \*\*\*\* or \*\*\*\* (a two-dimensional array of values of the specified `type`.)</span></span> |
| `parameters` | <span data-ttu-id="5c4b3-163">Tableau qui définit les paramètres d’entrée de la fonction.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-163">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="5c4b3-164">Les propriétés enfants `name` et `description` apparaissent dans l’intelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-164">The `name` and `description` child properties are used in the Excel intellisense.</span></span> <span data-ttu-id="5c4b3-165">Les propriétés enfants `type` et `dimensionality` sont identiques aux propriétés enfants de l’objet `result` décrit précédemment dans ce tableau.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-165">The `type` and `dimensionality` child properties are identical to the child properties of the `result` object that is described previously in this table.</span></span> |
| `options` | <span data-ttu-id="5c4b3-166">Vous permet de personnaliser certains aspects de la façon dont Excel exécute la fonction, et quand.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-166">The  property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="5c4b3-167">Pour plus d’informations sur l’utilisation de cette propriété, voir [Fonctions de flux](#streamed-functions) et [Annulation](#canceling-a-function) plus loin dans cet article.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-167">For more information about how this property can be used, see [Streamed functions](#streamed-functions) and [Cancellation](#canceling-a-function) later in this article.</span></span> |

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="5c4b3-168">Fonctions qui retournent des données provenant de sources externes</span><span class="sxs-lookup"><span data-stu-id="5c4b3-168">Functions that return data from external sources</span></span>

<span data-ttu-id="5c4b3-169">Si une fonction personnalisée récupère les données d’une source externe comme le Web, elle doit :</span><span class="sxs-lookup"><span data-stu-id="5c4b3-169">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="5c4b3-170">Renvoyer une promesse JavaScript à Excel.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-170">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="5c4b3-171">Résoudre la promesse avec la valeur finale en utilisant la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-171">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="5c4b3-172">Les fonctions personnalisées affichent un résultat temporaire `#GETTING_DATA` dans la cellule pendant qu’Excel attend le résultat final.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-172">Asynchronous functions display a `#GETTING_DATA` temporary error in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="5c4b3-173">Les utilisateurs peuvent interagir normalement avec le reste de la feuille de calcul tout en attendant le résultat.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-173">Users can interact normally with the rest of the spreadsheet while they wait for the result.</span></span>

<span data-ttu-id="5c4b3-174">Dans l’exemple de code suivant, la fonction personnalisée `getTemperature()` récupère la température actuelle d’un thermomètre.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-174">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer.</span></span> <span data-ttu-id="5c4b3-175">Notez que `sendWebRequest` est une fonction hypothétique, non spécifiée ici, qui utilise XHR pour appeler un service web de température.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-175">Note that `sendWebRequest` is a hypothetical function, not specified here, that uses XHR to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a><span data-ttu-id="5c4b3-176">Fonctions de flux</span><span class="sxs-lookup"><span data-stu-id="5c4b3-176">Streamed functions</span></span>

<span data-ttu-id="5c4b3-177">Les fonctions de flux personnalisées permettent de générer des données dans des cellules de manière répétée dans le temps, sans qu’un utilisateur doive demander explicitement le recalcul.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-177">Streamed custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request recalculation.</span></span> <span data-ttu-id="5c4b3-178">L’exemple de code suivant est une fonction personnalisée qui ajoute un nombre au résultat toutes les secondes.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-178">The following example is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="5c4b3-179">Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="5c4b3-179">Note the following about this code:</span></span>

- <span data-ttu-id="5c4b3-180">Excel affiche automatiquement chaque nouvelle valeur en utilisant le rappel `setResult`.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-180">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="5c4b3-181">Le dernier paramètre, `handler`, n’est jamais spécifié dans votre code d’enregistrement et ne s’affiche pas dans le menu de saisie semi-automatique pour les utilisateurs d’Excel lorsqu’ils lancent la fonction.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-181">For streamed functions, the final parameter, `handler`, is never specified in your registration code, and it does not display in the autocomplete menu to Excel users when they enter the function.</span></span> <span data-ttu-id="5c4b3-182">Il s’agit d’un objet contenant une fonction de rappel `setResult` utilisée pour transmettre des données de la fonction à Excel afin de mette à jour la valeur d’une cellule.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-182">It’s an object that contains a `setResult` callback function that’s used to pass data from the function to Excel to update the value of a cell.</span></span>

- <span data-ttu-id="5c4b3-183">Pour qu’Excel transmette la fonction `setResult` dans l'objet `handler`, vous devez déclarer la prise en charge de la diffusion en continu pendant l’enregistrement de votre fonction en définissant l’option `"stream": true` dans la propriété `options` pour la fonction personnalisée dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-183">In order for Excel to pass the `setResult` function in the `handler` object, you must declare support for streaming during your function registration by setting the option `"stream": true` in the `options` property for the custom function in the registration JSON file.</span></span>

```js
function incrementValue(increment, handler){
    var result = 0;
    setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);
}
```

## <a name="canceling-a-function"></a><span data-ttu-id="5c4b3-184">Annulation d’une fonction</span><span class="sxs-lookup"><span data-stu-id="5c4b3-184">Canceling a function</span></span>

<span data-ttu-id="5c4b3-185">Dans certains cas, vous devrez peut-être annuler l’exécution d’une fonction personnalisée en flux continu pour réduire la consommation de la bande passante, de la mémoire et de la charge processeur.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-185">In some situations, you may need to cancel the execution of a streamed custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="5c4b3-186">Excel annule l’exécution d’une fonction dans les situations suivantes :</span><span class="sxs-lookup"><span data-stu-id="5c4b3-186">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="5c4b3-187">Quand l’utilisateur modifie ou supprime une cellule qui fait référence à la fonction.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-187">The user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="5c4b3-188">Quand un des arguments (entrées) de la fonction est modifié.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-188">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="5c4b3-189">Dans ce cas, un nouvel appel de fonction est déclenché après l’annulation.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-189">In this case, a new function call is triggered in addition to the cancelation.</span></span>

- <span data-ttu-id="5c4b3-190">L’utilisateur déclenche manuellement un nouveau calcul.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-190">The user triggers recalculation manually.</span></span> <span data-ttu-id="5c4b3-191">Dans ce cas, un nouvel appel de fonction est déclenché après l’annulation.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-191">In this case, a new function call is triggered in addition to the cancelation.</span></span>

> [!NOTE]
> <span data-ttu-id="5c4b3-192">Vous devez implémenter un gestionnaire d'annulation pour chaque fonction de diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-192">You must implement a cancellation handler for every streaming function.</span></span>

<span data-ttu-id="5c4b3-193">Pour rendre une fonction annulable, définissez l’option `"cancelable": true` dans la propriété `options` pour la fonction personnalisée dans le fichier de métadonnées JSON.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-193">To make a function cancelable, set the option `"cancelable": true` in the `options` property for the custom function in the registration JSON file.</span></span>

<span data-ttu-id="5c4b3-194">Le code suivant affiche la même fonction `incrementValue` qui a été décrite précédemment, mais cette fois avec un gestionnaire d’annulation implémenté.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-194">The following code shows the same `incrementValue` function that was described previously, but this time with a cancellation handler implemented.</span></span> <span data-ttu-id="5c4b3-195">Dans cet exemple, `clearInterval()` s’exécute lorsque la fonction `incrementValue` est annulée.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-195">In this example, `clearInterval()` will run when the `incrementValue` function is canceled.</span></span>

```js
function incrementValue(increment, handler){
    var result = 0;
    var timer = setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);

    handler.onCanceled = function(){
        clearInterval(timer);
    }
}
```

## <a name="saving-and-sharing-state"></a><span data-ttu-id="5c4b3-196">Enregistrement et partage de l'état</span><span class="sxs-lookup"><span data-stu-id="5c4b3-196">Saving and sharing state</span></span>

<span data-ttu-id="5c4b3-197">Les fonctions personnalisées peuvent enregistrer des données dans des variables JavaScript globales.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-197">Custom functions can save data in global JavaScript variables.</span></span> <span data-ttu-id="5c4b3-198">Lors d’appels ultérieurs, votre fonction personnalisée pourra utiliser les valeurs enregistrées dans ces variables.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-198">In subsequent calls, your custom function may use the values saved in these variables.</span></span> <span data-ttu-id="5c4b3-199">L'état enregistré est utile lorsque les utilisateurs ajoutent la même fonction personnalisée à plusieurs cellules, car toutes les instances de la fonction peuvent partager l'état.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-199">Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state.</span></span> <span data-ttu-id="5c4b3-200">Par exemple, vous pouvez enregistrer les données renvoyées par un appel à une ressource web pour éviter d’effectuer des appels supplémentaires à la même ressource web.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-200">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="5c4b3-201">L’exemple de code suivant illustre une implémentation de la fonction de flux précédente relative à la température et qui enregistre l’état globalement.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-201">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="5c4b3-202">Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="5c4b3-202">Note the following about this code:</span></span>

- <span data-ttu-id="5c4b3-203">`refreshTemperature` ,est une fonction de flux qui chaque seconde, lit la température d’un thermomètre spécifique.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-203">`refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second.</span></span> <span data-ttu-id="5c4b3-204">Les nouvelles températures sont enregistrées dans la variable `savedTemperatures`, mais ne mettent pas directement à jour la valeur de la cellule.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-204">New temperatures are saved in the `savedTemperatures` variable, but does not directly update the cell value.</span></span> <span data-ttu-id="5c4b3-205">Elles ne doivent pas être appelées directement à partir d'une cellule de feuille de calcul, *de sorte qu'elles ne sont pas enregistrées dans le fichier JSON*.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-205">It should not be directly called from a worksheet cell, *so it is not registered in the JSON file*.</span></span>

- <span data-ttu-id="5c4b3-206">`streamTemperature` met à jour les valeurs de température affichées dans la cellule chaque seconde et utilise une variable `savedTemperatures` comme source de données.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-206">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span> <span data-ttu-id="5c4b3-207">Elles doivent être enregistrées dans le fichier JSON et nommées en lettres majuscules, `STREAMTEMPERATURE`.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-207">It must be registered in the JSON file, and named with all upper-case letters, `STREAMTEMPERATURE`.</span></span>

- <span data-ttu-id="5c4b3-208">Les utilisateurs peuvent appeler `streamTemperature` à partir de plusieurs cellules dans l’interface utilisateur Excel.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-208">Users may call `streamTemperature` from several cells in the Excel UI.</span></span> <span data-ttu-id="5c4b3-209">Chaque appel lit des données depuis la même variable `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-209">Each call reads data from the same `savedTemperatures` variable.</span></span>

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

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="5c4b3-210">Utilisation des plages de données</span><span class="sxs-lookup"><span data-stu-id="5c4b3-210">Working with ranges of data</span></span>

<span data-ttu-id="5c4b3-211">Votre fonction personnalisée peut accepter une plage de données comme paramètre d’entrée, ou elle peut renvoyer une plage de données.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-211">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="5c4b3-212">En JavaScript, une plage de données est représentée sous la forme d’un tableau à deux dimensions.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-212">In JavaScript, a range of data is represented as a 2-dimensional array.</span></span>

<span data-ttu-id="5c4b3-213">Par exemple, supposons que votre fonction renvoie la deuxième valeur la plus élevée prise dans une plage de nombres stockés dans Excel.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-213">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="5c4b3-214">La fonction suivante accepte le paramètre `values`, qui est de type `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-214">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="5c4b3-215">Notez que dans les métadonnées JSON de cette fonction, vous devez définir la propriété `type` du paramètre à `matrix`.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-215">Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="handling-errors"></a><span data-ttu-id="5c4b3-216">Gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="5c4b3-216">Handling errors</span></span>

<span data-ttu-id="5c4b3-217">Lorsque vous créez un complément qui définit des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-217">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="5c4b3-218">La gestion des erreurs pour les fonctions personnalisées est identique à la [gestion des erreurs pour l’API JavaScript d’Excel dans son ensemble](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="5c4b3-218">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="5c4b3-219">Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-219">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(x) {
    let url = "https://yourhypotheticalapi/comments/" + x;

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

## <a name="known-issues"></a><span data-ttu-id="5c4b3-220">Problèmes connus</span><span class="sxs-lookup"><span data-stu-id="5c4b3-220">Known issues</span></span>

- <span data-ttu-id="5c4b3-221">Les descriptions de paramètre et les URL d’aide ne sont pas encore utilisées par Excel.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-221">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="5c4b3-222">Les fonctions personnalisées ne sont actuellement pas disponibles sur Excel pour les clients mobiles.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-222">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="5c4b3-223">Les fonctions volatiles (celles qui recalculent automatiquement lorsque des modifications de données indépendantes sont effectuées dans la feuille de calcul) ne sont pas encore prises en charge.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-223">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="5c4b3-224">Le déploiement via le portail d'administration Office 365 et AppSource n'est pas encore activé.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-224">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="5c4b3-225">Les fonctions personnalisées dans Excel Online peuvent cesser de fonctionner pendant une session après une période d'inactivité.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-225">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="5c4b3-226">Actualisez la page du navigateur (F5) et entrez à nouveau une fonction personnalisée pour restaurer la fonction.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-226">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="5c4b3-227">Il est possible d’avoir le résultat temporaire **#GETTING_DATA** dans la ou les cellules d’une feuille de calcul si vous avez plusieurs compléments s’exécutant dans Microsoft Excel pour Windows.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-227">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows.</span></span> <span data-ttu-id="5c4b3-228">Fermez toutes les fenêtres Excel et redémarrez Excel.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-228">Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="5c4b3-229">Des outils de débogage spécifiques pour les fonctions personnalisées pourraient devenir disponibles à l’avenir.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-229">Debugging tools specifically for custom functions may be available in the future.</span></span> <span data-ttu-id="5c4b3-230">En attendant, vous pouvez déboguer sur Excel Online à l’aide des outils de développement F12.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-230">In the meantime, you can debug on Excel Online using F12 developer tools.</span></span> <span data-ttu-id="5c4b3-231">Voir plus de détails dans [Meilleures pratiques pour les fonctions personnalisées](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="5c4b3-231">See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>

## <a name="changelog"></a><span data-ttu-id="5c4b3-232">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="5c4b3-232">Changelog</span></span>

- <span data-ttu-id="5c4b3-233">**7 novembre 2017 :** mise à disposition\* de la préversion des fonctions personnalisées et d'exemples</span><span class="sxs-lookup"><span data-stu-id="5c4b3-233">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="5c4b3-234">**20 novembre 2017 :** correction du bogue de compatibilité pour les utilisateurs de la version 8801 et ultérieure</span><span class="sxs-lookup"><span data-stu-id="5c4b3-234">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="5c4b3-235">**28 novembre 2017 :** mise à disposition\* de la prise en charge de l’annulation sur des fonctions asynchrones (nécessite la modification des fonctions de flux)</span><span class="sxs-lookup"><span data-stu-id="5c4b3-235">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="5c4b3-236">**7 mai 2018**  : mise à disposition\* de la prise en charge pour Mac, Excel Online et fonctions synchrones en cours de traitement</span><span class="sxs-lookup"><span data-stu-id="5c4b3-236">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="5c4b3-237">**20 septembre 2018** : Support fourni pour les fonctions personnalisées à l'exécution de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="5c4b3-237">**September 20, 2018**: Shipped support for custom functions JavaScript runtime.</span></span> <span data-ttu-id="5c4b3-238">Pour plus d’informations, voir [Exécution des fonctions personnalisées d’Excel](custom-functions-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="5c4b3-238">For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>

<span data-ttu-id="5c4b3-239">\* vers le canal Office Insiders</span><span class="sxs-lookup"><span data-stu-id="5c4b3-239">\* to the Office Insiders Channel</span></span>

## <a name="see-also"></a><span data-ttu-id="5c4b3-240">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="5c4b3-240">See also</span></span>

* [<span data-ttu-id="5c4b3-241">Métadonnées des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="5c4b3-241">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="5c4b3-242">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="5c4b3-242">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="5c4b3-243">Meilleures pratiques pour les fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="5c4b3-243">Custom functions best practices</span></span>](custom-functions-best-practices.md)
