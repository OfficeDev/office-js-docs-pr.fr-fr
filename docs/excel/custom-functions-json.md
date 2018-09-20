# <a name="custom-function-metadata"></a><span data-ttu-id="49256-101">Métadonnées de fonction personnalisées</span><span class="sxs-lookup"><span data-stu-id="49256-101">Custom function metadata</span></span>

<span data-ttu-id="49256-102">Lorsque vous ajoutez des [fonctions personnalisées](custom-functions-overview.md) dans un complément Excel, vous devez héberger un fichier JSON qui contient des métadonnées sur les fonctions (en plus d'héberger un fichier JavaScript comportant des fonctions et un fichier HTML sans interface utilisateur devant servir de parent au fichier JavaScript).</span><span class="sxs-lookup"><span data-stu-id="49256-102">When you include [custom functions](custom-functions-overview.md) in an Excel add-in, you must host a JSON file that contains metadata about the functions (in addition to hosting a JavaScript file with the functions and a UI-less HTML file to serve as the parent of the JavaScript file).</span></span> <span data-ttu-id="49256-103">Cet article présente et illustre ce qu'est le format de fichier JSON.</span><span class="sxs-lookup"><span data-stu-id="49256-103">This article describes the format of the JSON file with examples.</span></span>

<span data-ttu-id="49256-104">Un échantillon de fichier JSON complet est disponible [ici](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="49256-104">A complete sample JSON file is available [here](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions-array"></a><span data-ttu-id="49256-105">Tableau de fonctions</span><span class="sxs-lookup"><span data-stu-id="49256-105">Functions array</span></span>

<span data-ttu-id="49256-106">Les métadonnées sont un objet JSON qui contient une seule `functions` propriété dont la valeur est un tableau d'objets.</span><span class="sxs-lookup"><span data-stu-id="49256-106">The metadata is a JSON object that contains a single `functions` property whose value is an array of objects.</span></span> <span data-ttu-id="49256-107">Chacun de ces objets représente une fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="49256-107">Each of these objects represents one custom function.</span></span> <span data-ttu-id="49256-108">Le tableau suivant contient ses propriétés :</span><span class="sxs-lookup"><span data-stu-id="49256-108">The following table contains its properties:</span></span>

|  <span data-ttu-id="49256-109">Propriété</span><span class="sxs-lookup"><span data-stu-id="49256-109">Property</span></span>  |  <span data-ttu-id="49256-110">Type de données</span><span class="sxs-lookup"><span data-stu-id="49256-110">Data Type</span></span>  |  <span data-ttu-id="49256-111">Obligatoire ?</span><span class="sxs-lookup"><span data-stu-id="49256-111">Required?</span></span>  |  <span data-ttu-id="49256-112">Description</span><span class="sxs-lookup"><span data-stu-id="49256-112">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="49256-113">chaîne</span><span class="sxs-lookup"><span data-stu-id="49256-113">string</span></span>  |  <span data-ttu-id="49256-114">Non</span><span class="sxs-lookup"><span data-stu-id="49256-114">No</span></span>  |  <span data-ttu-id="49256-115">Une description de la fonction figurant sur l'interface utilisateur Excel.</span><span class="sxs-lookup"><span data-stu-id="49256-115">A description of the function that appears in the Excel UI.</span></span> <span data-ttu-id="49256-116">Par exemple, " Convertir une valeur Celsius en Fahrenheit ".</span><span class="sxs-lookup"><span data-stu-id="49256-116">For example, "Converts a Celsius value to Fahrenheit".</span></span> |
|  `helpUrl`  |  <span data-ttu-id="49256-117">chaîne</span><span class="sxs-lookup"><span data-stu-id="49256-117">string</span></span>  |   <span data-ttu-id="49256-118">Non</span><span class="sxs-lookup"><span data-stu-id="49256-118">No</span></span>  |  <span data-ttu-id="49256-119">L’URL où vos utilisateurs peuvent obtenir de l’aide sur la fonction.</span><span class="sxs-lookup"><span data-stu-id="49256-119">URL where your users can get help about the function.</span></span> <span data-ttu-id="49256-120">(Il est affiché dans une tâche.) Par exemple, "http://contoso.com/help/convertcelsiustofahrenheit.html"</span><span class="sxs-lookup"><span data-stu-id="49256-120">(It is displayed in a taskpane.) For example, "http://contoso.com/help/convertcelsiustofahrenheit.html"</span></span>  |
|  `name`  |  <span data-ttu-id="49256-121">chaîne</span><span class="sxs-lookup"><span data-stu-id="49256-121">string</span></span>  |  <span data-ttu-id="49256-122">Oui</span><span class="sxs-lookup"><span data-stu-id="49256-122">Yes</span></span>  |  <span data-ttu-id="49256-123">Le nom de la fonction telle qu'elle apparaîtra (préfixée d'un espace de nom) dans l'interface utilisateur Excel lorsqu'un utilisateur sélectionne une fonction.</span><span class="sxs-lookup"><span data-stu-id="49256-123">The name of the function as it will appear (prepended with a namespace) in the Excel UI when a user is selecting a function.</span></span> <span data-ttu-id="49256-124">Il devrait être le même que le nom de la fonction où il est défini dans le JavaScript.</span><span class="sxs-lookup"><span data-stu-id="49256-124">It should be the same as the function's name where it is defined in the JavaScript.</span></span> |
|  `options`  |  <span data-ttu-id="49256-125">object</span><span class="sxs-lookup"><span data-stu-id="49256-125">object</span></span>  |  <span data-ttu-id="49256-126">Non</span><span class="sxs-lookup"><span data-stu-id="49256-126">No</span></span>  |  <span data-ttu-id="49256-127">Configurer comment Excel traite une fonction.</span><span class="sxs-lookup"><span data-stu-id="49256-127">Configure how Excel processes the function.</span></span> <span data-ttu-id="49256-128">Voir [options objet](#options-object) pour plus de détails.</span><span class="sxs-lookup"><span data-stu-id="49256-128">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="49256-129">array</span><span class="sxs-lookup"><span data-stu-id="49256-129">array</span></span>  |  <span data-ttu-id="49256-130">Oui</span><span class="sxs-lookup"><span data-stu-id="49256-130">Yes</span></span>  |  <span data-ttu-id="49256-131">Métadonnées sur les paramètres de la fonction.</span><span class="sxs-lookup"><span data-stu-id="49256-131">Metadata about the parameters to the function.</span></span> <span data-ttu-id="49256-132">Voir[tableau de paramètres](#parameters-array) pour plus de détails.</span><span class="sxs-lookup"><span data-stu-id="49256-132">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="49256-133">object</span><span class="sxs-lookup"><span data-stu-id="49256-133">object</span></span>  |  <span data-ttu-id="49256-134">Oui</span><span class="sxs-lookup"><span data-stu-id="49256-134">Yes</span></span>  |  <span data-ttu-id="49256-135">Métadonnées sur la valeur renvoyée par la fonction.</span><span class="sxs-lookup"><span data-stu-id="49256-135">Metadata about the value returned by the function.</span></span> <span data-ttu-id="49256-136">Voir [objet de résultat](#result-object) pour plus de détails.</span><span class="sxs-lookup"><span data-stu-id="49256-136">See [result object](#result-object) for details.</span></span> |

## <a name="options-object"></a><span data-ttu-id="49256-137">Objet Options</span><span class="sxs-lookup"><span data-stu-id="49256-137">Options object</span></span>

<span data-ttu-id="49256-138">L’ `options` objet configure comment Excel traite la fonction.</span><span class="sxs-lookup"><span data-stu-id="49256-138">The `options` object configures how Excel processes the function.</span></span> <span data-ttu-id="49256-139">Le tableau suivant contient ses propriétés :</span><span class="sxs-lookup"><span data-stu-id="49256-139">The following table contains its properties:</span></span>

|  <span data-ttu-id="49256-140">Propriété</span><span class="sxs-lookup"><span data-stu-id="49256-140">Property</span></span>  |  <span data-ttu-id="49256-141">Type de données</span><span class="sxs-lookup"><span data-stu-id="49256-141">Data Type</span></span>  |  <span data-ttu-id="49256-142">Obligatoire ?</span><span class="sxs-lookup"><span data-stu-id="49256-142">Required?</span></span>  |  <span data-ttu-id="49256-143">Description</span><span class="sxs-lookup"><span data-stu-id="49256-143">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="49256-144">booléen</span><span class="sxs-lookup"><span data-stu-id="49256-144">boolean</span></span>  |  <span data-ttu-id="49256-145">Non, la valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="49256-145">No, default is `false`.</span></span>  |  <span data-ttu-id="49256-p110">Si `true`, Excel appelle le gestionnaire `onCanceled` à chaque fois que l’utilisateur exécute une action qui a pour effet l’annulation de la fonction ; par exemple, déclencher manuellement le recalcul, ou modifier une cellule référencée par la fonction. Si vous utilisez cette option, Excel appelle la fonction JavaScript avec un paramètre `caller` en plus. (Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`). Dans le corps de la fonction, un gestionnaire doit être affecté au membre `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="49256-p110">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). In the body of the function, a handler must be assigned to the `caller.onCanceled` member. Note,  and  cannot both be .</span></span>|
|  `stream`  |  <span data-ttu-id="49256-150">booléen</span><span class="sxs-lookup"><span data-stu-id="49256-150">boolean</span></span>  |  <span data-ttu-id="49256-151">Non, la valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="49256-151">No, default is `false`.</span></span>  |  <span data-ttu-id="49256-152">Si `true`, la fonction peut générer une sortie plusieurs fois dans la cellule même lorsqu'elle n'est invoquée qu'une seule fois.</span><span class="sxs-lookup"><span data-stu-id="49256-152">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="49256-153">Cette option est utile pour les sources de données en évolution rapide, telles que le cours d'une action.</span><span class="sxs-lookup"><span data-stu-id="49256-153">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="49256-154">Si vous utilisez cette option, Excel appellera la fonction JavaScript avec un paramètre `caller` additionnel.</span><span class="sxs-lookup"><span data-stu-id="49256-154">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="49256-155">(Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`).</span><span class="sxs-lookup"><span data-stu-id="49256-155">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="49256-156">La fonction ne devrait pas avoir de `return` déclaration.</span><span class="sxs-lookup"><span data-stu-id="49256-156">The function should have no `return` statement.</span></span> <span data-ttu-id="49256-157">Au lieu de cela, la valeur du résultat est transmise en tant que motif de la `caller.setResult` méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="49256-157">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span>|

## <a name="parameters-array"></a><span data-ttu-id="49256-158">Tableau de paramètres</span><span class="sxs-lookup"><span data-stu-id="49256-158">Parameters array</span></span>

<span data-ttu-id="49256-159">La propriété `parameters`est un tableau d'objets.</span><span class="sxs-lookup"><span data-stu-id="49256-159">The `parameters` property is an array of objects.</span></span> <span data-ttu-id="49256-160">Chacun de ces objets représente un paramètre.</span><span class="sxs-lookup"><span data-stu-id="49256-160">Each of these objects represents a parameter.</span></span> <span data-ttu-id="49256-161">Le tableau suivant contient ses propriétés :</span><span class="sxs-lookup"><span data-stu-id="49256-161">The following table contains its properties:</span></span>

|  <span data-ttu-id="49256-162">Propriété</span><span class="sxs-lookup"><span data-stu-id="49256-162">Property</span></span>  |  <span data-ttu-id="49256-163">Type de données</span><span class="sxs-lookup"><span data-stu-id="49256-163">Data Type</span></span>  |  <span data-ttu-id="49256-164">Obligatoire ?</span><span class="sxs-lookup"><span data-stu-id="49256-164">Required?</span></span>  |  <span data-ttu-id="49256-165">Description</span><span class="sxs-lookup"><span data-stu-id="49256-165">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="49256-166">chaîne</span><span class="sxs-lookup"><span data-stu-id="49256-166">string</span></span>  |  <span data-ttu-id="49256-167">Non</span><span class="sxs-lookup"><span data-stu-id="49256-167">No</span></span> |  <span data-ttu-id="49256-168">Une description du paramètre.</span><span class="sxs-lookup"><span data-stu-id="49256-168">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="49256-169">chaîne</span><span class="sxs-lookup"><span data-stu-id="49256-169">string</span></span>  |  <span data-ttu-id="49256-170">Oui</span><span class="sxs-lookup"><span data-stu-id="49256-170">Yes</span></span>  |  <span data-ttu-id="49256-171">Doit être " scalaire ", ce qui signifie une valeur sans tableau, ou une " matrice ", ce qui signifie un tableau comportant des lignes.</span><span class="sxs-lookup"><span data-stu-id="49256-171">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `name`  |  <span data-ttu-id="49256-172">chaîne</span><span class="sxs-lookup"><span data-stu-id="49256-172">string</span></span>  |  <span data-ttu-id="49256-173">Oui</span><span class="sxs-lookup"><span data-stu-id="49256-173">Yes</span></span>  |  <span data-ttu-id="49256-174">Nom du paramètre.</span><span class="sxs-lookup"><span data-stu-id="49256-174">The name of the parameter.</span></span> <span data-ttu-id="49256-175">Ce nom est affiché dans IntelliSense d'Excel.</span><span class="sxs-lookup"><span data-stu-id="49256-175">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="49256-176">chaîne</span><span class="sxs-lookup"><span data-stu-id="49256-176">string</span></span>  |  <span data-ttu-id="49256-177">Oui</span><span class="sxs-lookup"><span data-stu-id="49256-177">Yes</span></span>  |  <span data-ttu-id="49256-178">Le type de données du paramètre.</span><span class="sxs-lookup"><span data-stu-id="49256-178">The data type of the parameter.</span></span> <span data-ttu-id="49256-179">Doit être " booléen ", " nombre " ou " chaîne ".</span><span class="sxs-lookup"><span data-stu-id="49256-179">Must be "boolean", "number", or "string".</span></span>  |

## <a name="result-object"></a><span data-ttu-id="49256-180">Objet de résultat</span><span class="sxs-lookup"><span data-stu-id="49256-180">Result object</span></span>

<span data-ttu-id="49256-181">La propriété `results`  fournit des métadonnées sur la valeur renvoyée par la fonction.</span><span class="sxs-lookup"><span data-stu-id="49256-181">The `results` property provides metadata about the value returned from the function.</span></span> <span data-ttu-id="49256-182">Le tableau suivant contient ses propriétés :</span><span class="sxs-lookup"><span data-stu-id="49256-182">The following table contains its properties:</span></span>

|  <span data-ttu-id="49256-183">Propriété</span><span class="sxs-lookup"><span data-stu-id="49256-183">Property</span></span>  |  <span data-ttu-id="49256-184">Type de données</span><span class="sxs-lookup"><span data-stu-id="49256-184">Data Type</span></span>  |  <span data-ttu-id="49256-185">Obligatoire ?</span><span class="sxs-lookup"><span data-stu-id="49256-185">Required?</span></span>  |  <span data-ttu-id="49256-186">Description</span><span class="sxs-lookup"><span data-stu-id="49256-186">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="49256-187">chaîne</span><span class="sxs-lookup"><span data-stu-id="49256-187">string</span></span>  |  <span data-ttu-id="49256-188">Non</span><span class="sxs-lookup"><span data-stu-id="49256-188">No</span></span>  |  <span data-ttu-id="49256-189">Doit être " scalaire ", ce qui signifie une valeur sans tableau, ou une " matrice ", ce qui signifie un tableau comportant des lignes.</span><span class="sxs-lookup"><span data-stu-id="49256-189">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `type`  |  <span data-ttu-id="49256-190">chaîne</span><span class="sxs-lookup"><span data-stu-id="49256-190">string</span></span>  |  <span data-ttu-id="49256-191">Oui</span><span class="sxs-lookup"><span data-stu-id="49256-191">Yes</span></span>  |  <span data-ttu-id="49256-192">Le type de données du paramètre.</span><span class="sxs-lookup"><span data-stu-id="49256-192">The data type of the parameter.</span></span> <span data-ttu-id="49256-193">Doit être " booléen ", " nombre " ou " chaîne ".</span><span class="sxs-lookup"><span data-stu-id="49256-193">Must be "boolean", "number", or "string".</span></span>  |

## <a name="example"></a><span data-ttu-id="49256-194">Exemple</span><span class="sxs-lookup"><span data-stu-id="49256-194">Example</span></span>

<span data-ttu-id="49256-195">Le code JSON suivant est un exemple de fichier de métadonnées pour fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="49256-195">The following JSON code is an example of a metadata file for custom functions.</span></span>

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
            ]
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
            ]
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
            ]
        },
        {
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": []
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
            ]
        }
    ]
}

```

## <a name="see-also"></a><span data-ttu-id="49256-196">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="49256-196">See also</span></span>
[<span data-ttu-id="49256-197">Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="49256-197">Custom functions</span></span>](custom-functions-overview.md)<br>
[<span data-ttu-id="49256-198">Directives et exemples de formules matricielles</span><span class="sxs-lookup"><span data-stu-id="49256-198">Guidelines and examples of array formulas</span></span>](https://support.office.com/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7)
