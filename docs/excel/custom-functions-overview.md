# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="340c1-101">Créer des fonctions personnalisées dans Excel (Aperçu)</span><span class="sxs-lookup"><span data-stu-id="340c1-101">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="340c1-102">Les fonctions personnalisées (similaires aux fonctions définies par l’utilisateur) permettent aux développeurs d’ajouter n’importe quelle fonction JavaScript à Excel en utilisant un complément.</span><span class="sxs-lookup"><span data-stu-id="340c1-102">Custom functions (similar to user-defined functions, or UDFs), allow developers to add any JavaScript function to Excel using an add-in.</span></span> <span data-ttu-id="340c1-103">Les utilisateurs peuvent alors avoir accès aux fonctions personnalisées comme toute autre fonction native dans Excel (telle que `=SUM()`).</span><span class="sxs-lookup"><span data-stu-id="340c1-103">Users can then access custom functions like any other native function in Excel (like =SUM()).</span></span> <span data-ttu-id="340c1-104">Cet article explique comment créer des fonctions personnalisées dans Excel.</span><span class="sxs-lookup"><span data-stu-id="340c1-104">This article explains how to create custom functions in Excel.</span></span>

<span data-ttu-id="340c1-105">L'illustration suivante montre comment un utilisateur final insère une fonction personnalisée dans une cellule.</span><span class="sxs-lookup"><span data-stu-id="340c1-105">The following illustration shows you how an end user would insert a custom function into a cell.</span></span> <span data-ttu-id="340c1-106">La fonction qui ajoute 42 à une paire de nombres.</span><span class="sxs-lookup"><span data-stu-id="340c1-106">Here’s the code for a sample custom function that adds 42 to a pair of numbers.</span></span>

<img alt="custom functions" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="340c1-107">Voici le code pour la même fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="340c1-107">Here’s the code for the same custom function.</span></span>

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

<span data-ttu-id="340c1-108">Les fonctions personnalisées sont désormais disponibles dans Developer Preview sous Windows, Mac et Excel Online.</span><span class="sxs-lookup"><span data-stu-id="340c1-108">Custom functions are now available in Developer Preview on Windows, Mac, and Excel Online.</span></span> <span data-ttu-id="340c1-109">Pour les tester, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="340c1-109">Follow these steps to try them:</span></span>

1.  <span data-ttu-id="340c1-110">Installez Office (version 9325 sur Windows ou 13.329 sur Mac) et participez au programme [Office Insider](https://products.office.com/en-us/office-insider).</span><span class="sxs-lookup"><span data-stu-id="340c1-110">Install Office (build 9325 on Windows or 13.329 on Mac) and join the [Office Insider](https://products.office.com/en-us/office-insider) program.</span></span> <span data-ttu-id="340c1-111">(Notez qu'il ne suffit pas d'obtenir la dernière version, la fonctionnalité sera désactivée sur n'importe quelle version jusqu'à ce que vous rejoignez le programme Insider)</span><span class="sxs-lookup"><span data-stu-id="340c1-111">(Note that it isn't enough just to get the latest build; the feature will be disabled on any build until you join the Insider program)</span></span>
2.  <span data-ttu-id="340c1-112">Clonez le dépôt des [fonctions Excel personnalisées](https://github.com/OfficeDev/Excel-Custom-Functions) et suivez les instructions dans le fichier README.md pour démarrer le complément dans Excel, apporter des modifications dans le code et déboguer.</span><span class="sxs-lookup"><span data-stu-id="340c1-112">Clone the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) repo and follow the instructions in the README.md to start the add-in in Excel, make changes in the code, and debug.</span></span>
3.  <span data-ttu-id="340c1-113">Saisissez `=CONTOSO.ADD42(1,2)` dans une cellule, puis appuyez sur **Entrée** pour exécuter la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="340c1-113">Type `=CONTOSO.ADD42(1,2)` into any cell, and press **Enter** to run the custom function.</span></span>

<span data-ttu-id="340c1-114">Reportez-vous à la section **Problèmes connus**à la fin de cet article qui inclut les limites actuelles des fonctions personnalisées et sera mise à jour au fil du temps.</span><span class="sxs-lookup"><span data-stu-id="340c1-114">See the Known Issues section at the end of this article, which includes current limitations of custom functions and will be updated over time.</span></span>

## <a name="learn-the-basics"></a><span data-ttu-id="340c1-115">Notions fondamentales</span><span class="sxs-lookup"><span data-stu-id="340c1-115">Learn the basics</span></span>

<span data-ttu-id="340c1-116">Dans le dépôt d’exemple cloné, vous trouverez les fichiers suivants :</span><span class="sxs-lookup"><span data-stu-id="340c1-116">In the cloned sample repo, you’ll see the following files:</span></span>

- <span data-ttu-id="340c1-117">**customfunctions.js**, qui contient le code de fonction personnalisé (voir l'exemple de code simple ci-dessus pour la fonction `ADD42`).</span><span class="sxs-lookup"><span data-stu-id="340c1-117">**customfunctions.js**, which contains the custom function code (see the simple code example above for the `ADD42` function).</span></span>
- <span data-ttu-id="340c1-118">**customfunctions.json**, qui contient l’enregistrement JSON qui indique à Excel votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="340c1-118">**customfunctions.json**, which contains the registration JSON that tells Excel about your custom function.</span></span> <span data-ttu-id="340c1-119">Avec l’enregistrement, vos fonctions personnalisées apparaissent dans la liste des fonctions disponibles affichée lorsqu'un utilisateur saisit du texte dans les cellules.</span><span class="sxs-lookup"><span data-stu-id="340c1-119">Registration makes your custom functions appear in the list of available functions displayed when users type in cells.</span></span>
- <span data-ttu-id="340c1-120">**customfunctions.html**, qui fournit une référence &lt;Scipt&gt; au fichier JS.</span><span class="sxs-lookup"><span data-stu-id="340c1-120">customfunctions.html, which provides a Script reference to customfunctions.js.</span></span> <span data-ttu-id="340c1-121">Ce fichier n’affiche pas d’interface utilisateur dans Excel.</span><span class="sxs-lookup"><span data-stu-id="340c1-121">This file does not display UI in Excel.</span></span>
- <span data-ttu-id="340c1-122">**customfunctions.xml**, qui indique à Excel l’emplacement des fichiers HTML, JavaScript et JSON, et spécifie également un espace de noms pour toutes les fonctions personnalisées installées avec le complément.</span><span class="sxs-lookup"><span data-stu-id="340c1-122">**customfunctions.xml**, which tells Excel the location of the HTML, JavaScript, and JSON files; and also specifies a namespace for all the custom functions that are installed with the add-in.</span></span>

### <a name="json-file-customfunctionsjson"></a><span data-ttu-id="340c1-123">Fichier JSON (customfunctions.json)</span><span class="sxs-lookup"><span data-stu-id="340c1-123">JSON file (customfunctions.json)</span></span>

<span data-ttu-id="340c1-124">Le code suivant dans customfunctions.json spécifie les métadonnées pour la même fonction `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="340c1-124">The following code in customfunctions.json specifies the metadata for the same `ADD42` function.</span></span>

> [!NOTE]
> <span data-ttu-id="340c1-125">Les informations de référence détaillées pour le fichier JSON, y compris les options non utilisées dans cet exemple, sont dans [Enregistrement des fonctions personnalisées JSON](https://dev.office.com/reference/add-ins/custom-functions-json).</span><span class="sxs-lookup"><span data-stu-id="340c1-125">Detailed reference information for the JSON file, including options not used in this example, is at [Custom Functions Registration JSON](https://dev.office.com/reference/add-ins/custom-functions-json).</span></span>

<span data-ttu-id="340c1-126">Notez que pour cet exemple :</span><span class="sxs-lookup"><span data-stu-id="340c1-126">Note that for this example:</span></span>

- <span data-ttu-id="340c1-127">Il n'y a qu'une seule fonction personnalisée, donc il n'y a qu'un seul membre d tableau `functions`.</span><span class="sxs-lookup"><span data-stu-id="340c1-127">There's only one custom function, so there's only one member of the `functions` array.</span></span>
- <span data-ttu-id="340c1-128">La propriété `name` définit le nom de la fonction.</span><span class="sxs-lookup"><span data-stu-id="340c1-128">The `name` property defines the function name.</span></span> <span data-ttu-id="340c1-129">Comme vous le voyez dans l'image gif animée montrée précédemment, un espace de noms (`CONTOSO`) est ajouté au nom de la fonction dans le menu remplissage automatique Excel.</span><span class="sxs-lookup"><span data-stu-id="340c1-129">As you see in the animated gif shown previously, a namespace (`CONTOSO`) is prepended to the function name in the Excel autocomplete menu.</span></span> <span data-ttu-id="340c1-130">Ce préfixe est défini dans le manifeste du complément, décrit ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="340c1-130">This prefix is defined in the add-in manifest, described below.</span></span> <span data-ttu-id="340c1-131">Le préfixe et le nom de la fonction sont séparés à l'aide d'un point et, par convention, les préfixes et les noms de fonctions sont en majuscules.</span><span class="sxs-lookup"><span data-stu-id="340c1-131">The prefix and the function name are separated using a period, and by convention prefixes and function names are uppercase.</span></span> <span data-ttu-id="340c1-132">Pour utiliser votre fonction personnalisée, un utilisateur tape l’espace de nom suivi du nom de la fonction (`ADD42`) dans une cellule, dans ce cas `=CONTOSO.ADD42`.</span><span class="sxs-lookup"><span data-stu-id="340c1-132">To use your custom function, a user types the namespace followed by the function's name (`ADD42`) into a cell, in this case `=CONTOSO.ADD42`.</span></span> <span data-ttu-id="340c1-133">Le préfixe est destiné à être utilisé comme identificateur de votre entreprise ou du complément.</span><span class="sxs-lookup"><span data-stu-id="340c1-133">The prefix is intended to be used as an identifier for your add-in.</span></span> 
- <span data-ttu-id="340c1-134">Le `description` apparaît dans le menu remplissage automatique dans Excel.</span><span class="sxs-lookup"><span data-stu-id="340c1-134">`description`: The description appears in the autocomplete menu in Excel.</span></span>
- <span data-ttu-id="340c1-135">Lorsque l’utilisateur demande de l’aide concernant une fonction, Excel ouvre un volet Office et affiche la page web accessible via cette URL spécifiée dans `helpUrl`.</span><span class="sxs-lookup"><span data-stu-id="340c1-135">`helpUrl`: When the user requests help for a function, Excel opens a task pane and displays the web page found at this URL.</span></span>
- <span data-ttu-id="340c1-136">La propriété `result` spécifie le type d’information retournée à Excel par la fonction.</span><span class="sxs-lookup"><span data-stu-id="340c1-136">`result`: Defines the type of information returned by the function to Excel.</span></span> <span data-ttu-id="340c1-137">La propriété enfant `type` peut `"string"`, `"number"`, ou `"boolean"`.</span><span class="sxs-lookup"><span data-stu-id="340c1-137">The `type` child property can `"string"`, `"number"`, or `"boolean"`.</span></span> <span data-ttu-id="340c1-138">La propriété `dimensionality` peut être `scalar` ou `matrix` (un tableau bidimensionnel de valeurs de la valeur spécifiée `type`).</span><span class="sxs-lookup"><span data-stu-id="340c1-138">The `dimensionality` property can be `scalar` or `matrix` (a two-dimensional array of values of the specified `type`.)</span></span>
- <span data-ttu-id="340c1-139">Le tableau `parameters` spécifie, *dans l'ordre*, le type de données dans chaque paramètre qui est passé à la fonction.</span><span class="sxs-lookup"><span data-stu-id="340c1-139">The `parameters` array specifies, *in order*, the type of data in each parameter that is passed to the function.</span></span> <span data-ttu-id="340c1-140">Les propriétés enfants `name` et `description` sont utilisées dans l’intelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="340c1-140">The `name` and `description` child properties are used in the Excel intellisense.</span></span> <span data-ttu-id="340c1-141">Les propriétés enfants `type` et `dimensionality` sont identiques aux propriétés enfants de la propriété `result` décrite ci-dessus.</span><span class="sxs-lookup"><span data-stu-id="340c1-141">The `type` and `dimensionality` child properties are identical to the child properties of the `result` property described above.</span></span>
- <span data-ttu-id="340c1-142">La propriété `options` vous permet de personnaliser certains aspects de la façon dont Excel exécute la fonction et quand.</span><span class="sxs-lookup"><span data-stu-id="340c1-142">The `options` property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="340c1-143">Il y a plus d'informations sur ces options plus loin dans cet article.</span><span class="sxs-lookup"><span data-stu-id="340c1-143">There is more information about these options later in this article.</span></span>

 ```js
{
    "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
    "functions": [
        {
            "name": "ADD42", 
            "description":  "adds 42 to the input numbers",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
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
            "options": {
                "sync": true
            }
        }
    ]
}
```

> [!NOTE]
> <span data-ttu-id="340c1-144">Les fonctions personnalisées sont enregistrées lorsqu’un utilisateur exécute le complément pour la première fois.</span><span class="sxs-lookup"><span data-stu-id="340c1-144">The custom functions are registered when a user runs the add-in for the first time.</span></span> <span data-ttu-id="340c1-145">Après cela, elles sont disponibles, pour le même utilisateur, dans tous les classeurs (pas seulement celui dans lequel le complément a fonctionné initialement.)</span><span class="sxs-lookup"><span data-stu-id="340c1-145">After that, they are available, for that same user, in all workbooks (not only the one where the add-in ran initially.)</span></span>

<span data-ttu-id="340c1-146">Vos paramètres de serveur pour le fichier JSON doivent avoir activé [CORS](https://developer.mozilla.org/en-US/docs/Web/HTTP/CORS) pour que les fonctions personnalisées fonctionnent correctement dans Excel Online.</span><span class="sxs-lookup"><span data-stu-id="340c1-146">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/en-US/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>


### <a name="manifest-file-customfunctionsxml"></a><span data-ttu-id="340c1-147">Fichier manifeste (customfunctions.xml)</span><span class="sxs-lookup"><span data-stu-id="340c1-147">Manifest file (customfunctions.xml)</span></span>


<span data-ttu-id="340c1-148">Ce qui suit est un exemple de balisage `<ExtensionPoint>` et `<Resources>` à inclure dans le manifeste du complément pour permettre à Excel d’exécuter vos fonctions.</span><span class="sxs-lookup"><span data-stu-id="340c1-148">The following is an example of the `<ExtensionPoint>` and `<Resources>` markup that you include in the add-in's manifest to enable Excel to run your functions.</span></span> <span data-ttu-id="340c1-149">Notez ce qui suit à propos de ce balisage :</span><span class="sxs-lookup"><span data-stu-id="340c1-149">Note the following facts about this markup:</span></span>

- <span data-ttu-id="340c1-150">L’élément `<Script>` et son ID de ressources correspondante spécifie l’emplacement du fichier JavaScript avec vos fonctions.</span><span class="sxs-lookup"><span data-stu-id="340c1-150">The `<Script>` element and its corresponding resource ID specifies the location of the JavaScript file with your functions.</span></span>
- <span data-ttu-id="340c1-151">L'élément `<Page>` et son ID de ressources correspondante spécifie l'emplacement de la page HTML de votre complément.</span><span class="sxs-lookup"><span data-stu-id="340c1-151">The `<Page>` element and its corresponding resource ID specifies the location of the HTML page of your add-in.</span></span> <span data-ttu-id="340c1-152">La page HTML comprend un tag `<Script>` qui charge le fichier JavaScript (customfunctions.js).</span><span class="sxs-lookup"><span data-stu-id="340c1-152">The HTML page includes a `<Script>` tag that loads the JavaScript file (customfunctions.js).</span></span> <span data-ttu-id="340c1-153">La page HTML est une page masquée qui n’est jamais affichée dans l’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="340c1-153">The HTML page is a hidden page and is never displayed in the UI.</span></span>
- <span data-ttu-id="340c1-154">L’élément `<Metadata>` et son ID de ressources correspondante spécifie l’emplacement du fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="340c1-154">The `<Metadata>` element and its corresponding resource ID specifies the location of the JSON file.</span></span>
- <span data-ttu-id="340c1-155">Un élément `<Namespace>` et son ID de ressources correspondante spécifie le préfixe pour toutes les fonctions personnalisées dans le complément.</span><span class="sxs-lookup"><span data-stu-id="340c1-155">A `<Namespace>` element and its corresponding resource ID specifies the prefix for all custom functions in the add-in.</span></span>


```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1\_0">
    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="residjs" />
                    </Script>
                    <Page>
                        <SourceLocation resid="residhtml"/>
                    </Page>
                    <Metadata>
                        <SourceLocation resid="residjson" />
                    </Metadata>
                    <Namespace resid="residNS" />
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>
    <Resources>
        <bt:Urls>
            <bt:Url id="residjson" DefaultValue="http://127.0.0.1:8080/customfunctions.json" />
            <bt:Url id="residjs" DefaultValue="http://127.0.0.1:8080/customfunctions.js" />
            <bt:Url id="residhtml" DefaultValue="http://127.0.0.1:8080/customfunctions.html" />
        </bt:Urls>
        <bt:ShortStrings>
            <bt:String id="residNS" DefaultValue="CONTOSO" />
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>

```

## <a name="initializing-custom-functions"></a><span data-ttu-id="340c1-156">Initialisation des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="340c1-156">Initializing custom functions</span></span>

<span data-ttu-id="340c1-157">Votre code doit initialiser la fonctionnalité de fonctions personnalisées avant de l'utiliser.</span><span class="sxs-lookup"><span data-stu-id="340c1-157">Your code must initialize the custom functions feature before using it.</span></span> <span data-ttu-id="340c1-158">Vous pouvez le faire soit dans un tag &lt;Script&gt; dans le fichier HTML (customfunctions.html) ou en haut du fichier JavaScript (customfunctions.js).</span><span class="sxs-lookup"><span data-stu-id="340c1-158">You can do this either in a &lt;Script&gt; tag in the HTML file (customfunctions.html) or at the top of the JavaScript file (customfunctions.js).</span></span> <span data-ttu-id="340c1-159">Lors de l'aperçu des fonctions personnalisées, vous avez le choix entre deux syntaxes pour l'initialisation.</span><span class="sxs-lookup"><span data-stu-id="340c1-159">During the preview of custom functions, you have your choice of two syntaxes for intializing.</span></span> <span data-ttu-id="340c1-160">Le fichier HTML dans le référentiel utilise la syntaxe suivante :</span><span class="sxs-lookup"><span data-stu-id="340c1-160">The HTML file in the repo uses the following syntax:</span></span>

```js
Office.initialize = function (reason) {
    return Excel.CustomFunctions.initialize();
};
```

<span data-ttu-id="340c1-161">Vous pouvez également utiliser la syntaxe suivante :</span><span class="sxs-lookup"><span data-stu-id="340c1-161">You can also use the following syntax:</span></span>

```js
Office.Preview.StartCustomFunctions();
```

## <a name="synchronous-and-asynchronous-functions"></a><span data-ttu-id="340c1-162">Fonctions synchrones et asynchrones</span><span class="sxs-lookup"><span data-stu-id="340c1-162">Synchronous and asynchronous functions</span></span>

<span data-ttu-id="340c1-163">La fonction `ADD42` ci-dessus est synchrone par rapport à Excel (désigné en réglant les paramètres de l'option `"sync": true` dans le fichier JSON).</span><span class="sxs-lookup"><span data-stu-id="340c1-163">The function `ADD42` above is synchronous with respect to Excel (designated by setting the option `"sync": true` in the JSON file).</span></span> <span data-ttu-id="340c1-164">Les fonctions synchrones offrent des performances rapides car elles s’exécutent dans le même processus qu’Excel et s’exécutent en parallèle lors du calcul multithread.</span><span class="sxs-lookup"><span data-stu-id="340c1-164">Synchronous functions offer fast performance because they run in the same process as Excel and they run in parallel during multithreaded calculation.</span></span>   

<span data-ttu-id="340c1-165">D'un autre côté, si votre fonction personnalisée récupère des données du Web, elle doit être asynchrone par rapport à Excel.</span><span class="sxs-lookup"><span data-stu-id="340c1-165">On the other hand, if your custom function retrieves data from the web, it must be asynchronous with respect to Excel.</span></span> <span data-ttu-id="340c1-166">Les fonctions asynchrones doivent :</span><span class="sxs-lookup"><span data-stu-id="340c1-166">Asynchronous functions must:</span></span>

1. <span data-ttu-id="340c1-167">Renvoyer une promesse JavaScript à Excel.</span><span class="sxs-lookup"><span data-stu-id="340c1-167">Return a JavaScript Promise to Excel.</span></span>
3. <span data-ttu-id="340c1-168">Résolvez la promesse avec la valeur finale en utilisant la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="340c1-168">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="340c1-169">Le code suivant indique un exemple de fonction personnalisée asynchrone qui récupère la température d’un thermomètre.</span><span class="sxs-lookup"><span data-stu-id="340c1-169">The following code shows an example of a custom function that retrieves the temperature of a thermometer.</span></span> <span data-ttu-id="340c1-170">Notez que `sendWebRequest` est une fonction hypothétique, non spécifiée ici, qui utilise XHR pour appeler un service Web de température.</span><span class="sxs-lookup"><span data-stu-id="340c1-170">Note that `sendWebRequest` is a hypothetical function, not specified here, that uses XHR to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

<span data-ttu-id="340c1-171">Les fonctions asynchrones affichent une erreur temporaire `GETTING_DATA` dans la cellule pendant qu'Excel attend le résultat final.</span><span class="sxs-lookup"><span data-stu-id="340c1-171">Asynchronous functions display a `GETTING_DATA` temporary error in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="340c1-172">Les utilisateurs peuvent interagir normalement avec le reste du tableur pendant qu’ils attendent le résultat.</span><span class="sxs-lookup"><span data-stu-id="340c1-172">Users can interact normally with the rest of the spreadsheet while they wait for the result.</span></span>

> [!NOTE]
> <span data-ttu-id="340c1-173">Les fonctions personnalisées sont asynchrones par défaut.</span><span class="sxs-lookup"><span data-stu-id="340c1-173">Custom functions are asynchronous by default.</span></span> <span data-ttu-id="340c1-174">Pour désigner les fonctions comme synchrones, définissez l’option `"sync": true` dans la propriété `options` pour la fonction personnalisée dans le fichier JSON d’enregistrement.</span><span class="sxs-lookup"><span data-stu-id="340c1-174">To designate functions as synchronous set the option `"sync": true` in the `options` property for the custom function in the registration JSON file.</span></span>

## <a name="streamed-functions"></a><span data-ttu-id="340c1-175">Fonctions de flux</span><span class="sxs-lookup"><span data-stu-id="340c1-175">Streamed functions</span></span>

<span data-ttu-id="340c1-176">Une fonction asynchrone peut être diffusée.</span><span class="sxs-lookup"><span data-stu-id="340c1-176">An asynchronous function can be streamed.</span></span> <span data-ttu-id="340c1-177">Les fonctions personnalisées de flux vous permettent d’afficher des données dans des cellules à plusieurs reprises au fil du temps, sans devoir attendre qu’Excel ou que des utilisateurs demandent à effectuer le calcul à nouveau.</span><span class="sxs-lookup"><span data-stu-id="340c1-177">Streamed custom functions let you output data to cells repeatedly over time, without waiting for Excel or users to request recalculations.</span></span> <span data-ttu-id="340c1-178">L’exemple suivant est une fonction personnalisée qui ajoute un nombre au résultat toutes les secondes.</span><span class="sxs-lookup"><span data-stu-id="340c1-178">The following example is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="340c1-179">Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="340c1-179">Note the following about this code:</span></span>

- <span data-ttu-id="340c1-180">Excel affiche automatiquement chaque nouvelle valeur en utilisant le rappel `setResult`.</span><span class="sxs-lookup"><span data-stu-id="340c1-180">Excel displays each new value automatically using the `setResult` callback.</span></span>
- <span data-ttu-id="340c1-181">Le paramètre final, `caller`, n’est jamais spécifié dans votre code d’enregistrement et ne s’affiche pas dans le menu de remplissage automatique pour les utilisateurs d’Excel lorsqu’ils entrent la fonction.</span><span class="sxs-lookup"><span data-stu-id="340c1-181">For streamed functions, the final parameter, `caller`, is never specified in your registration code, and it does not display in the autocomplete menu to Excel users when they enter the function.</span></span> <span data-ttu-id="340c1-182">Il s’agit d’un objet contenant une fonction de rappel `setResult` utilisée pour transmettre des données de la fonction à Excel afin de mette à jour la valeur d’une cellule.</span><span class="sxs-lookup"><span data-stu-id="340c1-182">It’s an object that contains a `setResult` callback function that’s used to pass data from the function to Excel to update the value of a cell.</span></span>
- <span data-ttu-id="340c1-183">Pour qu'Excel transmette la fonction `setResult` dans l'objet `caller`, vous devez déclarer la prise en charge de la diffusion en continu pendant l’enregistrement de votre fonction en définissant l’option `"stream": true` dans la propriété `options` pour la fonction personnalisée dans le fichier JSON d’enregistrement.</span><span class="sxs-lookup"><span data-stu-id="340c1-183">In order for Excel to pass the `setResult` function in the `caller` object, you must declare support for streaming during your function registration by setting the option `"stream": true` in the `options` property for the custom function in the registration JSON file.</span></span>

```js
function incrementValue(increment, caller){
    var result = 0;
    setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);
}
```

## <a name="cancellation"></a><span data-ttu-id="340c1-184">Annulation</span><span class="sxs-lookup"><span data-stu-id="340c1-184">Cancellation</span></span>

<span data-ttu-id="340c1-185">Vous pouvez annuler les fonctions de flux et les fonctions asynchrones.</span><span class="sxs-lookup"><span data-stu-id="340c1-185">You can cancel streamed functions and asynchronous functions.</span></span> <span data-ttu-id="340c1-186">L’annulation de vos appels de fonction permet de considérablement réduire leur consommation de bande passante, la mémoire de travail et la charge de l’UC.</span><span class="sxs-lookup"><span data-stu-id="340c1-186">Canceling your function calls is important to reduce their bandwith consumption, working memory, and CPU load.</span></span> <span data-ttu-id="340c1-187">Excel annule les appels de fonction dans les situations suivantes :</span><span class="sxs-lookup"><span data-stu-id="340c1-187">Excel cancels function calls in the following situations:</span></span>

- <span data-ttu-id="340c1-188">L’utilisateur modifie ou supprime une cellule qui fait référence à la fonction.</span><span class="sxs-lookup"><span data-stu-id="340c1-188">The user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="340c1-189">Un des arguments (entrées) de la fonction est modifié.</span><span class="sxs-lookup"><span data-stu-id="340c1-189">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="340c1-190">Dans ce cas, un nouvel appel de fonction est déclenché en plus de l’annulation.</span><span class="sxs-lookup"><span data-stu-id="340c1-190">In this case, a new function call is triggered in addition to the cancelation.</span></span>
- <span data-ttu-id="340c1-p124">L’utilisateur déclenche le nouveau processus de calcul manuellement. Comme pour le cas précédent, un nouvel appel de fonction est déclenché en plus de l’annulation.</span><span class="sxs-lookup"><span data-stu-id="340c1-p124">The user triggers recalculation manually. As with the above case, a new function call is triggered in addition to the cancelation.</span></span>

<span data-ttu-id="340c1-193">Vous *devez* implémenter un gestionnaire d'annulation pour chaque fonction de diffusion en continu.</span><span class="sxs-lookup"><span data-stu-id="340c1-193">You *must* implement a cancellation handler for every streaming function.</span></span> <span data-ttu-id="340c1-194">Les fonctions asynchrones, non diffusées en continu peuvent être annulables ou non, c'est à vous de décider.</span><span class="sxs-lookup"><span data-stu-id="340c1-194">Asynchronous, non-streaming functions may or may not be cancelable; it's up to you.</span></span> <span data-ttu-id="340c1-195">Les fonctions synchrones ne peuvent pas être annulées.</span><span class="sxs-lookup"><span data-stu-id="340c1-195">Synchronous functions cannot be canceled.</span></span>

<span data-ttu-id="340c1-196">Pour rendre une fonction annulable, définissez l’option `"cancelable": true` dans la propriété `options` pour la fonction personnalisée dans le fichier JSON d’enregistrement.</span><span class="sxs-lookup"><span data-stu-id="340c1-196">To make a function cancelable, set the option `"cancelable": true` in the `options` property for the custom function in the registration JSON file.</span></span>

<span data-ttu-id="340c1-197">Le code suivant affiche l’exemple précédent avec l’annulation mise en œuvre.</span><span class="sxs-lookup"><span data-stu-id="340c1-197">The following code shows the previous example with cancellation implemented.</span></span> <span data-ttu-id="340c1-198">Dans le code, l’objet `caller` contient une fonction `onCanceled` qui doit être définie pour chaque fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="340c1-198">In the code, the `caller` object contains an `onCanceled` function which should be defined for each custom function.</span></span>

```js
function incrementValue(increment, caller){ 
    var result = 0;
    var timer = setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);

    caller.onCanceled = function(){
        clearInterval(timer);
    }
}
```

## <a name="saving-and-sharing-state"></a><span data-ttu-id="340c1-199">Enregistrement et partage de l'état</span><span class="sxs-lookup"><span data-stu-id="340c1-199">Saving and sharing state</span></span>

<span data-ttu-id="340c1-200">Les fonctions asynchrones peuvent enregistrer des données dans des variables JavaScript globales.</span><span class="sxs-lookup"><span data-stu-id="340c1-200">Custom functions can save data in global JavaScript variables.</span></span> <span data-ttu-id="340c1-201">Lors d’appels ultérieurs, votre fonction personnalisée peut utiliser les valeurs enregistrées dans ces variables.</span><span class="sxs-lookup"><span data-stu-id="340c1-201">In subsequent calls, your custom function may use the values saved in these variables.</span></span> <span data-ttu-id="340c1-202">L'état enregistré est utile lorsque les utilisateurs ajoutent la même fonction personnalisée à plusieurs cellules, car toutes les instances de la fonction peuvent partager l'état.</span><span class="sxs-lookup"><span data-stu-id="340c1-202">Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state.</span></span> <span data-ttu-id="340c1-203">Par exemple, vous pouvez enregistrer les données renvoyées par un appel à une ressource web pour éviter d’effectuer des appels supplémentaires à la même ressource web.</span><span class="sxs-lookup"><span data-stu-id="340c1-203">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="340c1-204">Le code suivant illustre une implémentation de la fonction de diffusion en continu précédente relative à la température qui enregistre l’état à l’aide la variable.</span><span class="sxs-lookup"><span data-stu-id="340c1-204">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="340c1-205">Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="340c1-205">Note the following about this code:</span></span>

- <span data-ttu-id="340c1-206">`refreshTemperature` est une fonction diffusée en continu qui lit la température d’un thermomètre spécifique à chaque seconde qui passe.</span><span class="sxs-lookup"><span data-stu-id="340c1-206">`refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second.</span></span> <span data-ttu-id="340c1-207">Les nouvelles températures sont enregistrées dans la variable `savedTemperatures`, mais ne mettent pas directement à jour la valeur de la cellule.</span><span class="sxs-lookup"><span data-stu-id="340c1-207">New temperatures are saved in the `savedTemperatures` variable, but does not directly update the cell value.</span></span> <span data-ttu-id="340c1-208">Elles ne doivent pas être appelées directement à partir d'une cellule de feuille de calcul, *de sorte qu'elles ne sont pas enregistrées dans le fichier JSON*.</span><span class="sxs-lookup"><span data-stu-id="340c1-208">It should not be directly called from a worksheet cell, *so it is not registered in the JSON file*.</span></span>
- <span data-ttu-id="340c1-209">`streamTemperature` met à jour les valeurs de température affichées dans la cellule chaque seconde et utilise une variable `savedTemperatures` comme source de données.</span><span class="sxs-lookup"><span data-stu-id="340c1-209">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span> <span data-ttu-id="340c1-210">Elles doivent être enregistrées dans le fichier JSON et nommées en lettres majuscules, `STREAMTEMPERATURE`.</span><span class="sxs-lookup"><span data-stu-id="340c1-210">It must be registered in the JSON file, and named with all upper-case letters, `STREAMTEMPERATURE`.</span></span>
- <span data-ttu-id="340c1-211">Les utilisateurs peuvent appeler `streamTemperature` à partir de plusieurs cellules dans l’interface utilisateur Excel.</span><span class="sxs-lookup"><span data-stu-id="340c1-211">Users may call `streamTemperature` from several cells in the Excel UI.</span></span> <span data-ttu-id="340c1-212">Chaque appel lit des données de la même variable `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="340c1-212">Each call reads data from the same `savedTemperatures` variable.</span></span>

```js
var savedTemperatures;

function streamTemperature(thermometerID, caller){ 
     if(!savedTemperatures[thermometerID]){
         refreshTemperatures(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
     }

     function getNextTemperature(){
         caller.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
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

> [!NOTE]
> <span data-ttu-id="340c1-213">Les fonctions synchrones (désignées en paramétrant l'option `"sync": true` dans le fichier JSON) ne peuvent pas partager l'état car Excel les parallélise lors du calcul multithread.</span><span class="sxs-lookup"><span data-stu-id="340c1-213">Synchronous functions (designated by setting the option `"sync": true` in the JSON file) cannot share state because Excel parallelizes them during multithreaded calculation.</span></span> <span data-ttu-id="340c1-214">Seules les fonctions asynchrones peuvent partager l'état car les fonctions synchrones d'un complément partagent le même contexte JavaScript dans chaque session.</span><span class="sxs-lookup"><span data-stu-id="340c1-214">Only asynchronous functions may share state because an add-in's synchronous functions share the same JavaScript context in each session.</span></span>

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="340c1-215">Utilisation des plages de données</span><span class="sxs-lookup"><span data-stu-id="340c1-215">Working with ranges of data</span></span>

<span data-ttu-id="340c1-216">Votre fonction personnalisée accepte les plages de données en tant que paramètres. Sinon, vous pouvez renvoyer une plage de données à partir d’une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="340c1-216">Your custom function can take a range of data as a parameter, or you can return a range of data from a custom function.</span></span>

<span data-ttu-id="340c1-217">Par exemple, supposons que votre fonction renvoie la deuxième valeur la plus élevée parmi une plage de nombres stockée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="340c1-217">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="340c1-218">La fonction suivante prend le paramètre `values`, c’est-à-dire un type de paramètre `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="340c1-218">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="340c1-219">Notez que dans l'enregistrement JSON pour cette fonction, vous devez définir le paramètre propriété `type`sur `matrix`.</span><span class="sxs-lookup"><span data-stu-id="340c1-219">Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.</span></span>

```js
function secondHighest(values){ 
     var highest = values[0][0], secondHighest = values[0][0];
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

<span data-ttu-id="340c1-220">Comme vous pouvez le voir, les plages sont gérées en JavaScript sous la forme de tableaux de tableaux de lignes (comme un tableau à deux dimensions).</span><span class="sxs-lookup"><span data-stu-id="340c1-220">As you can see, ranges are handled in JavaScript as arrays of row arrays (like a 2-dimensional array).</span></span>

## <a name="known-issues"></a><span data-ttu-id="340c1-221">Problèmes connus</span><span class="sxs-lookup"><span data-stu-id="340c1-221">Known issues</span></span>

- <span data-ttu-id="340c1-222">Les descriptions de paramètre et les URL d’aide ne sont pas encore utilisées par Excel.</span><span class="sxs-lookup"><span data-stu-id="340c1-222">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="340c1-223">Les fonctions personnalisées ne sont actuellement pas disponibles sur Excel pour les clients mobiles.</span><span class="sxs-lookup"><span data-stu-id="340c1-223">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="340c1-224">Actuellement, les compléments s’appuient sur un processus de navigateur masqué pour exécuter les fonctions asynchrones.</span><span class="sxs-lookup"><span data-stu-id="340c1-224">Currently, add-ins rely on a hidden browser process to run custom functions.</span></span> <span data-ttu-id="340c1-225">À l’avenir, JavaScript s’exécutera directement sur certaines plateformes pour garantir que les fonctions personnalisées sont plus rapides et utilisent moins de mémoire.</span><span class="sxs-lookup"><span data-stu-id="340c1-225">In the future, JavaScript will run directly on some platforms to ensure custom functions are faster and use less memory.</span></span> <span data-ttu-id="340c1-226">Par ailleurs, la page HTML référencée par l’élément `<Page>`dans le manifeste ne sera pas nécessaire pour la plupart des plateformes, car Excel exécutera directement le code JavaScript.</span><span class="sxs-lookup"><span data-stu-id="340c1-226">Additionally, the HTML page referenced by the `<Page>`Page element in the manifest won’t be needed for most platforms because Excel will run the JavaScript directly.</span></span> <span data-ttu-id="340c1-227">Pour vous préparer à ce changement, vérifiez que vos fonctions personnalisées n’utilisent pas le DOM de page web.</span><span class="sxs-lookup"><span data-stu-id="340c1-227">To prepare for this change, ensure your custom functions do not use the webpage DOM.</span></span> <span data-ttu-id="340c1-228">Les API hôtes prises en charge pour accéder au Web seront [WebSocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) et [XHR](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) en utilisant GET ou POST.</span><span class="sxs-lookup"><span data-stu-id="340c1-228">The supported host APIs for accessing the web will be [WebSocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) and [XHR](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) using GET or POST.</span></span>
- <span data-ttu-id="340c1-229">Les fonctions volatiles (celles qui recalculent automatiquement lorsque des modifications de données indépendantes sont effectuées dans le tableur) ne sont pas encore prises en charge.</span><span class="sxs-lookup"><span data-stu-id="340c1-229">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="340c1-230">Le débogage est uniquement activé pour les fonctions asynchrones sur Excel pour Windows.</span><span class="sxs-lookup"><span data-stu-id="340c1-230">Debugging is only enabled for asynchronous functions on Excel for Windows.</span></span>
- <span data-ttu-id="340c1-231">Le déploiement via le portail d'administration Office 365 et AppSource n'est pas encore activé.</span><span class="sxs-lookup"><span data-stu-id="340c1-231">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="340c1-232">Les fonctions personnalisées dans Excel Online peuvent cesser de fonctionner pendant une session après une période d'inactivité.</span><span class="sxs-lookup"><span data-stu-id="340c1-232">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="340c1-233">Actualisez la page du navigateur (F5) et entrez à nouveau une fonction personnalisée pour restaurer la fonction.</span><span class="sxs-lookup"><span data-stu-id="340c1-233">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>

## <a name="changelog"></a><span data-ttu-id="340c1-234">Journal des modifications</span><span class="sxs-lookup"><span data-stu-id="340c1-234">Changelog</span></span>

- <span data-ttu-id="340c1-235">**7 novembre 2017 :** mise à disposition des exemples et de la version d’évaluation des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="340c1-235">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="340c1-236">**20 novembre 2017 :** correction du bogue de compatibilité pour les utilisateurs de la version 8801 et ultérieure</span><span class="sxs-lookup"><span data-stu-id="340c1-236">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="340c1-237">**28 novembre 2017 :** prise en charge de l’annulation sur des fonctions asynchrones (nécessite la modification des fonctions de flux)</span><span class="sxs-lookup"><span data-stu-id="340c1-237">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="340c1-238">**7 mai 2018**: Support fourni pour Mac, Excel Online et fonctions synchrones en cours de traitement</span><span class="sxs-lookup"><span data-stu-id="340c1-238">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
