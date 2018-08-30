# <a name="custom-function-metadata"></a><span data-ttu-id="71740-101">Métadonnées de fonction personnalisées</span><span class="sxs-lookup"><span data-stu-id="71740-101">Custom function metadata</span></span>

<span data-ttu-id="71740-102">Lorsque vous ajoutez des [fonctions personnalisées](custom-functions-overview.md) dans un complément Excel, vous devez héberger un fichier JSON qui contient des métadonnées sur les fonctions (en plus d'héberger un fichier JavaScript comportant des fonctions et un fichier HTML sans interface utilisateur devant servir de parent au fichier JavaScript).</span><span class="sxs-lookup"><span data-stu-id="71740-102">When you include [custom functions](custom-functions-overview.md) in an Excel add-in, you must host a JSON file that contains metadata about the functions (in addition to hosting a JavaScript file with the functions and a UI-less HTML file to serve as the parent of the JavaScript file).</span></span> <span data-ttu-id="71740-103">Cet article présente et illustre ce qu'est le format de fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="71740-103">This article describes the format of the JSON file with examples.</span></span>

<span data-ttu-id="71740-104">Un échantillon de fichier JSON complet est disponible [ici](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="71740-104">A complete sample JSON file is available [here](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions-array"></a><span data-ttu-id="71740-105">Tableau de fonctions</span><span class="sxs-lookup"><span data-stu-id="71740-105">Functions array</span></span>

<span data-ttu-id="71740-106">Les métadonnées sont un objet JSON qui contient une seule `functions` propriété dont la valeur est un tableau d'objets.</span><span class="sxs-lookup"><span data-stu-id="71740-106">The metadata is a JSON object that contains a single `functions` property whose value is an array of objects.</span></span> <span data-ttu-id="71740-107">Chacun de ces objets représente une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="71740-107">Each of these objects represents one custom function.</span></span> <span data-ttu-id="71740-108">Le tableau suivant contient ses propriétés :</span><span class="sxs-lookup"><span data-stu-id="71740-108">The following table contains its properties:</span></span>

|  <span data-ttu-id="71740-109">Propriété</span><span class="sxs-lookup"><span data-stu-id="71740-109">Property</span></span>  |  <span data-ttu-id="71740-110">Type de données</span><span class="sxs-lookup"><span data-stu-id="71740-110">Data Type</span></span>  |  <span data-ttu-id="71740-111">Obligatoire ?</span><span class="sxs-lookup"><span data-stu-id="71740-111">Required?</span></span>  |  <span data-ttu-id="71740-112">Description</span><span class="sxs-lookup"><span data-stu-id="71740-112">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="71740-113">chaîne</span><span class="sxs-lookup"><span data-stu-id="71740-113">string</span></span>  |  <span data-ttu-id="71740-114">Non</span><span class="sxs-lookup"><span data-stu-id="71740-114">No</span></span>  |  <span data-ttu-id="71740-115">Une description de la fonction figurant sur l'interface utilisateur Excel.</span><span class="sxs-lookup"><span data-stu-id="71740-115">A description of the function that appears in the Excel UI.</span></span> <span data-ttu-id="71740-116">Par exemple, " Convertir une valeur Celsius en Fahrenheit ".</span><span class="sxs-lookup"><span data-stu-id="71740-116">For example, "Converts a Celsius value to Fahrenheit".</span></span> |
|  `helpUrl`  |  <span data-ttu-id="71740-117">chaîne</span><span class="sxs-lookup"><span data-stu-id="71740-117">string</span></span>  |   <span data-ttu-id="71740-118">Non</span><span class="sxs-lookup"><span data-stu-id="71740-118">No</span></span>  |  <span data-ttu-id="71740-119">L’URL où vos utilisateurs peuvent obtenir de l’aide sur la fonction.</span><span class="sxs-lookup"><span data-stu-id="71740-119">URL where your users can get help about the function.</span></span> <span data-ttu-id="71740-120">(Il est affiché dans une tâche.) Par exemple, "http://contoso.com/help/convertcelsiustofahrenheit.html"</span><span class="sxs-lookup"><span data-stu-id="71740-120">(It is displayed in a taskpane.) For example, "http://contoso.com/help/convertcelsiustofahrenheit.html"</span></span>  |
|  `name`  |  <span data-ttu-id="71740-121">chaîne</span><span class="sxs-lookup"><span data-stu-id="71740-121">string</span></span>  |  <span data-ttu-id="71740-122">Oui</span><span class="sxs-lookup"><span data-stu-id="71740-122">Yes</span></span>  |  <span data-ttu-id="71740-123">Le nom de la fonction telle qu'elle apparaîtra (préfixée d'un espace de nom) dans l'interface utilisateur Excel lorsqu'un utilisateur sélectionne une fonction.</span><span class="sxs-lookup"><span data-stu-id="71740-123">The name of the function as it will appear (prepended with a namespace) in the Excel UI when a user is selecting a function.</span></span> <span data-ttu-id="71740-124">Il devrait être le même que le nom de la fonction où il est défini dans le JavaScript.</span><span class="sxs-lookup"><span data-stu-id="71740-124">It should be the same as the function's name where it is defined in the JavaScript.</span></span> |
|  `options`  |  <span data-ttu-id="71740-125">objet</span><span class="sxs-lookup"><span data-stu-id="71740-125">object</span></span>  |  <span data-ttu-id="71740-126">Non</span><span class="sxs-lookup"><span data-stu-id="71740-126">No</span></span>  |  <span data-ttu-id="71740-127">Configurer comment Excel traite une fonction.</span><span class="sxs-lookup"><span data-stu-id="71740-127">Configure how Excel processes the function.</span></span> <span data-ttu-id="71740-128">Voir [options objet](#options-object) pour plus de détails.</span><span class="sxs-lookup"><span data-stu-id="71740-128">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="71740-129">tableau</span><span class="sxs-lookup"><span data-stu-id="71740-129">array</span></span>  |  <span data-ttu-id="71740-130">Oui</span><span class="sxs-lookup"><span data-stu-id="71740-130">Yes</span></span>  |  <span data-ttu-id="71740-131">Métadonnées sur les paramètres de la fonction.</span><span class="sxs-lookup"><span data-stu-id="71740-131">Metadata about the parameters to the function.</span></span> <span data-ttu-id="71740-132">Voir[tableau de paramètres](#parameters-array) pour plus de détails.</span><span class="sxs-lookup"><span data-stu-id="71740-132">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="71740-133">objet</span><span class="sxs-lookup"><span data-stu-id="71740-133">object</span></span>  |  <span data-ttu-id="71740-134">Oui</span><span class="sxs-lookup"><span data-stu-id="71740-134">Yes</span></span>  |  <span data-ttu-id="71740-135">Métadonnées sur la valeur renvoyée par la fonction.</span><span class="sxs-lookup"><span data-stu-id="71740-135">Metadata about the value returned by the function.</span></span> <span data-ttu-id="71740-136">Voir [objet de résultat](#result-object) pour plus de détails.</span><span class="sxs-lookup"><span data-stu-id="71740-136">See [result object](#result-object) for details.</span></span> |

## <a name="options-object"></a><span data-ttu-id="71740-137">Objet Options</span><span class="sxs-lookup"><span data-stu-id="71740-137">Options object</span></span>

<span data-ttu-id="71740-138">L’ `options` objet configure comment Excel traite la fonction.</span><span class="sxs-lookup"><span data-stu-id="71740-138">The `options` object configures how Excel processes the function.</span></span> <span data-ttu-id="71740-139">Le tableau suivant contient ses propriétés :</span><span class="sxs-lookup"><span data-stu-id="71740-139">The following table contains its properties:</span></span>

|  <span data-ttu-id="71740-140">Propriété</span><span class="sxs-lookup"><span data-stu-id="71740-140">Property</span></span>  |  <span data-ttu-id="71740-141">Type de données</span><span class="sxs-lookup"><span data-stu-id="71740-141">Data Type</span></span>  |  <span data-ttu-id="71740-142">Obligatoire ?</span><span class="sxs-lookup"><span data-stu-id="71740-142">Required?</span></span>  |  <span data-ttu-id="71740-143">Description</span><span class="sxs-lookup"><span data-stu-id="71740-143">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="71740-144">booléen</span><span class="sxs-lookup"><span data-stu-id="71740-144">boolean</span></span>  |  <span data-ttu-id="71740-145">Non, la valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="71740-145">No, default is `false`.</span></span>  |  <span data-ttu-id="71740-146">Lorsqu’`true`Excel appelle le `onCanceled` gestionnaire au moment où l'utilisateur prend une action visant par exemple à annuler la fonction, le déclenchement manuel du recalcul ou la modification d’une cellule est référencée par cette fonction.</span><span class="sxs-lookup"><span data-stu-id="71740-146">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="71740-147">Si vous utilisez cette option, Excel appellera la fonction JavaScript avec un paramètre `caller` additionnel.</span><span class="sxs-lookup"><span data-stu-id="71740-147">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="71740-148">(Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`).</span><span class="sxs-lookup"><span data-stu-id="71740-148">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="71740-149">Dans le corps de la fonction, un gestionnaire doit être affecté à un membre `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="71740-149">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="71740-150">Remarque : `cancelable` et `sync` ne peuvent pas être à la fois `true`.</span><span class="sxs-lookup"><span data-stu-id="71740-150">Note, `cancelable` and `sync` cannot both be `true`.</span></span>  |
|  `stream`  |  <span data-ttu-id="71740-151">booléen</span><span class="sxs-lookup"><span data-stu-id="71740-151">boolean</span></span>  |  <span data-ttu-id="71740-152">Non, la valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="71740-152">No, default is `false`.</span></span>  |  <span data-ttu-id="71740-153">Si `true`, la fonction peut générer une sortie plusieurs fois dans la cellule même lorsqu'elle n'est invoquée qu'une seule fois.</span><span class="sxs-lookup"><span data-stu-id="71740-153">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="71740-154">Cette option est utile pour les sources de données en évolution rapide, telles que le cours d'une action.</span><span class="sxs-lookup"><span data-stu-id="71740-154">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="71740-155">Si vous utilisez cette option, Excel appellera la fonction JavaScript avec un paramètre `caller` additionnel.</span><span class="sxs-lookup"><span data-stu-id="71740-155">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="71740-156">(Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`).</span><span class="sxs-lookup"><span data-stu-id="71740-156">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="71740-157">La fonction ne devrait pas avoir de `return` déclaration.</span><span class="sxs-lookup"><span data-stu-id="71740-157">The function should have no `return` statement.</span></span> <span data-ttu-id="71740-158">Au lieu de cela, la valeur du résultat est transmise en tant que motif de la `caller.setResult` méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="71740-158">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="71740-159">Remarque : `stream` et `sync` ne peuvent pas être à la fois `true`.</span><span class="sxs-lookup"><span data-stu-id="71740-159">Note, `stream` and `sync` may not both be `true`.</span></span>|
|  `sync`  |  <span data-ttu-id="71740-160">booléen</span><span class="sxs-lookup"><span data-stu-id="71740-160">boolean</span></span>  |  <span data-ttu-id="71740-161">Non, la valeur par défaut est `false`</span><span class="sxs-lookup"><span data-stu-id="71740-161">No, default is `false`</span></span>  |  <span data-ttu-id="71740-162">Si `true`, la fonction s'exécute de manière synchrone et elle doit renvoyer une valeur.</span><span class="sxs-lookup"><span data-stu-id="71740-162">If `true`, the function runs synchronously and it must return a value.</span></span> <span data-ttu-id="71740-163">Si `false`, la fonction s'exécute de manière asynchrone et elle doit renvoyer un `OfficeExtension.Promise` objet.</span><span class="sxs-lookup"><span data-stu-id="71740-163">If `false`, the function runs asynchronously and it must return a `OfficeExtension.Promise` object.</span></span> <span data-ttu-id="71740-164">Remarque : `sync` n'est peut être pas `true` si `cancelable` ou `stream` sont `true`.</span><span class="sxs-lookup"><span data-stu-id="71740-164">Note, `sync`  may not be `true` if either `cancelable` or `stream` are `true`.</span></span>  |

## <a name="parameters-array"></a><span data-ttu-id="71740-165">Tableau de paramètres</span><span class="sxs-lookup"><span data-stu-id="71740-165">Parameters array</span></span>

<span data-ttu-id="71740-166">La propriété `parameters`est un tableau d'objets.</span><span class="sxs-lookup"><span data-stu-id="71740-166">The `parameters` property is an array of objects.</span></span> <span data-ttu-id="71740-167">Chacun de ces objets représente un paramètre.</span><span class="sxs-lookup"><span data-stu-id="71740-167">Each of these objects represents a parameter.</span></span> <span data-ttu-id="71740-168">Le tableau suivant contient ses propriétés :</span><span class="sxs-lookup"><span data-stu-id="71740-168">The following table contains its properties:</span></span>

|  <span data-ttu-id="71740-169">Propriété</span><span class="sxs-lookup"><span data-stu-id="71740-169">Property</span></span>  |  <span data-ttu-id="71740-170">Type de données</span><span class="sxs-lookup"><span data-stu-id="71740-170">Data Type</span></span>  |  <span data-ttu-id="71740-171">Obligatoire ?</span><span class="sxs-lookup"><span data-stu-id="71740-171">Required?</span></span>  |  <span data-ttu-id="71740-172">Description</span><span class="sxs-lookup"><span data-stu-id="71740-172">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="71740-173">chaîne</span><span class="sxs-lookup"><span data-stu-id="71740-173">string</span></span>  |  <span data-ttu-id="71740-174">Non</span><span class="sxs-lookup"><span data-stu-id="71740-174">No</span></span> |  <span data-ttu-id="71740-175">Une description du paramètre.</span><span class="sxs-lookup"><span data-stu-id="71740-175">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="71740-176">chaîne</span><span class="sxs-lookup"><span data-stu-id="71740-176">string</span></span>  |  <span data-ttu-id="71740-177">Oui</span><span class="sxs-lookup"><span data-stu-id="71740-177">Yes</span></span>  |  <span data-ttu-id="71740-178">Doit être " scalaire ", ce qui signifie une valeur sans tableau, ou une " matrice ", ce qui signifie un tableau comportant des lignes.</span><span class="sxs-lookup"><span data-stu-id="71740-178">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `name`  |  <span data-ttu-id="71740-179">chaîne</span><span class="sxs-lookup"><span data-stu-id="71740-179">string</span></span>  |  <span data-ttu-id="71740-180">Oui</span><span class="sxs-lookup"><span data-stu-id="71740-180">Yes</span></span>  |  <span data-ttu-id="71740-181">Nom du paramètre.</span><span class="sxs-lookup"><span data-stu-id="71740-181">The name of the parameter.</span></span> <span data-ttu-id="71740-182">Ce nom est affiché dans IntelliSense d'Excel.</span><span class="sxs-lookup"><span data-stu-id="71740-182">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="71740-183">chaîne</span><span class="sxs-lookup"><span data-stu-id="71740-183">string</span></span>  |  <span data-ttu-id="71740-184">Oui</span><span class="sxs-lookup"><span data-stu-id="71740-184">Yes</span></span>  |  <span data-ttu-id="71740-185">Le type de données du paramètre.</span><span class="sxs-lookup"><span data-stu-id="71740-185">The data type of the parameter.</span></span> <span data-ttu-id="71740-186">Doit être " booléen ", " nombre " ou " chaîne ".</span><span class="sxs-lookup"><span data-stu-id="71740-186">Must be "boolean", "number", or "string".</span></span>  |

## <a name="result-object"></a><span data-ttu-id="71740-187">Objet de résultat</span><span class="sxs-lookup"><span data-stu-id="71740-187">Result object</span></span>

<span data-ttu-id="71740-188">La propriété `results`  fournit des métadonnées sur la valeur renvoyée par la fonction.</span><span class="sxs-lookup"><span data-stu-id="71740-188">The `results` property provides metadata about the value returned from the function.</span></span> <span data-ttu-id="71740-189">Le tableau suivant contient ses propriétés :</span><span class="sxs-lookup"><span data-stu-id="71740-189">The following table contains its properties:</span></span>

|  <span data-ttu-id="71740-190">Propriété</span><span class="sxs-lookup"><span data-stu-id="71740-190">Property</span></span>  |  <span data-ttu-id="71740-191">Type de données</span><span class="sxs-lookup"><span data-stu-id="71740-191">Data Type</span></span>  |  <span data-ttu-id="71740-192">Obligatoire ?</span><span class="sxs-lookup"><span data-stu-id="71740-192">Required?</span></span>  |  <span data-ttu-id="71740-193">Description</span><span class="sxs-lookup"><span data-stu-id="71740-193">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="71740-194">chaîne</span><span class="sxs-lookup"><span data-stu-id="71740-194">string</span></span>  |  <span data-ttu-id="71740-195">Non</span><span class="sxs-lookup"><span data-stu-id="71740-195">No</span></span>  |  <span data-ttu-id="71740-196">Doit être " scalaire ", ce qui signifie une valeur sans tableau, ou une " matrice ", ce qui signifie un tableau comportant des lignes.</span><span class="sxs-lookup"><span data-stu-id="71740-196">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `type`  |  <span data-ttu-id="71740-197">chaîne</span><span class="sxs-lookup"><span data-stu-id="71740-197">string</span></span>  |  <span data-ttu-id="71740-198">Oui</span><span class="sxs-lookup"><span data-stu-id="71740-198">Yes</span></span>  |  <span data-ttu-id="71740-199">Le type de données du paramètre.</span><span class="sxs-lookup"><span data-stu-id="71740-199">The data type of the parameter.</span></span> <span data-ttu-id="71740-200">Doit être " booléen ", " nombre " ou " chaîne ".</span><span class="sxs-lookup"><span data-stu-id="71740-200">Must be "boolean", "number", or "string".</span></span>  |

## <a name="example"></a><span data-ttu-id="71740-201">Exemple</span><span class="sxs-lookup"><span data-stu-id="71740-201">Example</span></span>

<span data-ttu-id="71740-202">Le code JSON suivant est un exemple de fichier de métadonnées pour fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="71740-202">The following JSON code is an example of a metadata file for custom functions.</span></span>

```json
{
    "functions": [
        {
            "name": "ADD42", 
            "description":  "Adds 42 to the input number",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ],
            "options": {
                "sync": true
            }
        },
        {
            "name": "ADD42ASYNC", 
            "description":  "asynchronously wait 250ms, then add 42",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ],
            "options": {
                "sync": false
            }
        },
        {
            "name": "ISEVEN", 
            "description":  "Determines whether a number is even",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "boolean",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "the number to be evaluated",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ],
            "options": {
                "sync": true
            }
        },
        {
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": [],
            "options": {
                "sync": true
            }
        },
        {
            "name": "INCREMENTVALUE", 
            "description":  "Counts up from zero",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
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
                "sync": false,
                "stream": true,
                "cancelable": true
            }
        },
        {
            "name": "SECONDHIGHEST", 
            "description":  "gets the second highest number from a range",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "range",
                    "description": "the input range",
                    "type": "number",
                    "dimensionality": "matrix"
                }
            ],
            "options": {
                "sync": true
            }
        }
    ]
}

```

## <a name="see-also"></a><span data-ttu-id="71740-203">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="71740-203">See also</span></span>
[<span data-ttu-id="71740-204">Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="71740-204">Custom functions</span></span>](custom-functions-overview.md)<br>
[<span data-ttu-id="71740-205">Directives et exemples de formules matricielles</span><span class="sxs-lookup"><span data-stu-id="71740-205">Guidelines and examples of array formulas</span></span>](https://support.office.com/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7)
