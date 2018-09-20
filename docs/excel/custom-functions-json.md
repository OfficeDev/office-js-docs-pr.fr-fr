# <a name="custom-function-metadata"></a>Métadonnées de fonction personnalisées

Lorsque vous ajoutez des [fonctions personnalisées](custom-functions-overview.md) dans un complément Excel, vous devez héberger un fichier JSON qui contient des métadonnées sur les fonctions (en plus d'héberger un fichier JavaScript comportant des fonctions et un fichier HTML sans interface utilisateur devant servir de parent au fichier JavaScript). Cet article présente et illustre ce qu'est le format de fichier JSON.

Un échantillon de fichier JSON complet est disponible [ici](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).

## <a name="functions-array"></a>Tableau de fonctions

Les métadonnées sont un objet JSON qui contient une seule `functions` propriété dont la valeur est un tableau d'objets. Chacun de ces objets représente une fonction personnalisée. Le tableau suivant contient ses propriétés :

|  Propriété  |  Type de données  |  Obligatoire ?  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  chaîne  |  Non  |  Une description de la fonction figurant sur l'interface utilisateur Excel. Par exemple, " Convertir une valeur Celsius en Fahrenheit ". |
|  `helpUrl`  |  chaîne  |   Non  |  L’URL où vos utilisateurs peuvent obtenir de l’aide sur la fonction. (Il est affiché dans une tâche.) Par exemple, "http://contoso.com/help/convertcelsiustofahrenheit.html"  |
|  `name`  |  chaîne  |  Oui  |  Le nom de la fonction telle qu'elle apparaîtra (préfixée d'un espace de nom) dans l'interface utilisateur Excel lorsqu'un utilisateur sélectionne une fonction. Il devrait être le même que le nom de la fonction où il est défini dans le JavaScript. |
|  `options`  |  object  |  Non  |  Configurer comment Excel traite une fonction. Voir [options objet](#options-object) pour plus de détails. |
|  `parameters`  |  array  |  Oui  |  Métadonnées sur les paramètres de la fonction. Voir[tableau de paramètres](#parameters-array) pour plus de détails. |
|  `result`  |  object  |  Oui  |  Métadonnées sur la valeur renvoyée par la fonction. Voir [objet de résultat](#result-object) pour plus de détails. |

## <a name="options-object"></a>Objet Options

L’ `options` objet configure comment Excel traite la fonction. Le tableau suivant contient ses propriétés :

|  Propriété  |  Type de données  |  Obligatoire ?  |  Description  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  booléen  |  Non, la valeur par défaut est `false`.  |  Si `true`, Excel appelle le gestionnaire `onCanceled` à chaque fois que l’utilisateur exécute une action qui a pour effet l’annulation de la fonction ; par exemple, déclencher manuellement le recalcul, ou modifier une cellule référencée par la fonction. Si vous utilisez cette option, Excel appelle la fonction JavaScript avec un paramètre `caller` en plus. (Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`). Dans le corps de la fonction, un gestionnaire doit être affecté au membre `caller.onCanceled`.|
|  `stream`  |  booléen  |  Non, la valeur par défaut est `false`.  |  Si `true`, la fonction peut générer une sortie plusieurs fois dans la cellule même lorsqu'elle n'est invoquée qu'une seule fois. Cette option est utile pour les sources de données en évolution rapide, telles que le cours d'une action. Si vous utilisez cette option, Excel appellera la fonction JavaScript avec un paramètre `caller` additionnel. (Ne ***pas*** enregistrer ce paramètre dans la propriété `parameters`). La fonction ne devrait pas avoir de `return` déclaration. Au lieu de cela, la valeur du résultat est transmise en tant que motif de la `caller.setResult` méthode de rappel.|

## <a name="parameters-array"></a>Tableau de paramètres

La propriété `parameters`est un tableau d'objets. Chacun de ces objets représente un paramètre. Le tableau suivant contient ses propriétés :

|  Propriété  |  Type de données  |  Obligatoire ?  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  chaîne  |  Non |  Une description du paramètre.  |
|  `dimensionality`  |  chaîne  |  Oui  |  Doit être " scalaire ", ce qui signifie une valeur sans tableau, ou une " matrice ", ce qui signifie un tableau comportant des lignes.  |
|  `name`  |  chaîne  |  Oui  |  Nom du paramètre. Ce nom est affiché dans IntelliSense d'Excel.  |
|  `type`  |  chaîne  |  Oui  |  Le type de données du paramètre. Doit être " booléen ", " nombre " ou " chaîne ".  |

## <a name="result-object"></a>Objet de résultat

La propriété `results`  fournit des métadonnées sur la valeur renvoyée par la fonction. Le tableau suivant contient ses propriétés :

|  Propriété  |  Type de données  |  Obligatoire ?  |  Description  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  chaîne  |  Non  |  Doit être " scalaire ", ce qui signifie une valeur sans tableau, ou une " matrice ", ce qui signifie un tableau comportant des lignes.  |
|  `type`  |  chaîne  |  Oui  |  Le type de données du paramètre. Doit être " booléen ", " nombre " ou " chaîne ".  |

## <a name="example"></a>Exemple

Le code JSON suivant est un exemple de fichier de métadonnées pour fonctions personnalisées.

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

## <a name="see-also"></a>Voir aussi
[Fonctions personnalisées](custom-functions-overview.md)<br>
[Directives et exemples de formules matricielles](https://support.office.com/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7)
