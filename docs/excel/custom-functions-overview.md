# <a name="create-custom-functions-in-excel-preview"></a>Créer des fonctions personnalisées dans Excel (Aperçu)

Les fonctions personnalisées (similaires aux fonctions définies par l’utilisateur, ou UDF) permettent aux développeurs d’ajouter n’importe quelle fonction JavaScript à Excel en utilisant un complément. Les utilisateurs peuvent alors avoir accès aux fonctions personnalisées comme toute autre fonction native dans Excel (par exemple, `=SUM()`). Cet article explique comment créer des fonctions personnalisées dans Excel.

L'illustration suivante montre comment un utilisateur final insère une fonction personnalisée dans une cellule. La fonction qui ajoute 42 à une paire de nombres.

<img alt="custom functions" src="../images/custom-function.gif" width="579" height="383" />

Voici le code pour la même fonction personnalisée.

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

Les fonctions personnalisées sont désormais disponibles dans Developer Preview sous Windows, Mac et Excel Online. Pour les tester, procédez comme suit :

1. Installez Office (version 9325 sur Windows ou 13.329 sur Mac) et participez au programme [Office Insider](https://products.office.com/office-insider). (Notez qu'il ne suffit pas d'obtenir la dernière version, la fonctionnalité sera désactivée sur n'importe quelle version jusqu'à ce que vous rejoignez le programme Insider)
2. Créez un projet de complément de fonctions personnalisées Excel à l’aide de [Yo Office](https://github.com/OfficeDev/generator-office)et suivez les instructions fournies dans le [projet README.md](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) pour démarrer le complément dans Excel, apporter des modifications dans le code et déboguer.
3. Saisissez `=CONTOSO.ADD42(1,2)` dans une cellule, puis appuyez sur **Entrée** pour exécuter la fonction personnalisée.

Reportez-vous à la section **Problèmes connus** à la fin de cet article qui inclut les limites actuelles des fonctions personnalisées et qui sera mise à jour au fil du temps.

## <a name="learn-the-basics"></a>Notions fondamentales

Dans le référentiel d’exemple cloné, vous trouverez les fichiers suivants :

- **./src/customfunctions.js**, qui contient le code de fonction personnalisée (voir l’exemple de code simple ci-dessus pour la fonction `ADD42`).
- **customfunctions.json**, qui contient l’enregistrement JSON qui indique à Excel votre fonction personnalisée. Avec l’enregistrement, vos fonctions personnalisées apparaissent dans la liste des fonctions disponibles qui s'affiche lorsqu'un utilisateur saisit du texte dans une cellule.
- **customfunctions.html**, qui fournit une référence &lt;Script&gt; au fichier JS. Ce fichier n’affiche pas d’interface utilisateur dans Excel.
- **./manifest.xml**, qui indique à Excel l’emplacement des fichiers HTML, JavaScript et JSON, et spécifie également un espace de noms pour toutes les fonctions personnalisées installées avec le complément.

### <a name="json-file-configcustomfunctionsjson"></a>Fichier JSON (./config/customfunctions.json)

Le code suivant dans customfunctions.json spécifie les métadonnées pour la même fonction `ADD42`.

> [!NOTE]
> Les informations de référence détaillées pour le fichier JSON, y compris les options non utilisées dans cet exemple, sont dans [Enregistrement des fonctions personnalisées JSON](custom-functions-json.md).

Notez que pour cet exemple :

- Il n'y a qu'une seule fonction personnalisée, donc il n'y a qu'un seul membre d tableau `functions`.
- La propriété `name` définit le nom de la fonction. Comme vous le voyez dans l'image gif animée montrée précédemment, un espace de noms (`CONTOSO`) est ajouté au nom de la fonction dans le menu remplissage automatique Excel. Ce préfixe est défini dans le manifeste du complément, décrit ci-dessous. Le préfixe et le nom de la fonction sont séparés à l'aide d'un point et, par convention, les préfixes et les noms de fonctions sont en majuscules. Pour utiliser votre fonction personnalisée, un utilisateur tape l’espace de nom suivi du nom de la fonction (`ADD42`) dans une cellule, dans ce cas `=CONTOSO.ADD42`. Le préfixe est destiné à être utilisé comme identificateur pour votre entreprise ou votre complément. 
- Le `description` apparaît dans le menu remplissage automatique dans Excel.
- Lorsque l’utilisateur demande de l’aide concernant une fonction, Excel ouvre un volet Office et affiche la page web accessible via cette URL spécifiée dans `helpUrl`.
- La propriété `result` spécifie le type d’informations renvoyées par la fonction à Excel. La propriété enfant `type` peut `"string"`, `"number"`, ou `"boolean"`. La propriété `dimensionality` peut être `scalar` ou `matrix` (un tableau bidimensionnel de valeurs de la valeur spécifiée `type`).
- Le tableau `parameters` spécifie, *dans l'ordre*, le type de données dans chaque paramètre qui est passé à la fonction. Les propriétés enfants `name` et `description` sont utilisées dans l’intelliSense Excel. Les propriétés enfants `type` et `dimensionality` sont identiques aux propriétés enfants de la propriété `result` décrite ci-dessus.
- La propriété `options` vous permet de personnaliser certains aspects de la façon dont Excel exécute la fonction et quand. Il y a plus d'informations sur ces options plus loin dans cet article.

```js
    {
        "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
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
> Les fonctions personnalisées sont enregistrées lorsqu’un utilisateur exécute le complément pour la première fois. Après cela, elles sont disponibles, pour le même utilisateur, dans tous les classeurs (pas seulement celui dans lequel le complément a fonctionné initialement.)

Vos paramètres de serveur pour le fichier JSON doivent avoir activé [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) pour que les fonctions personnalisées fonctionnent correctement dans Excel Online.


### <a name="manifest-file-manifestxml"></a>Fichier manifeste (./manifest.xml)


Ce qui suit est un exemple de balisage `<ExtensionPoint>` et `<Resources>` à inclure dans le manifeste du complément pour permettre à Excel d’exécuter vos fonctions. Notez ce qui suit à propos de ce balisage :

- L’élément `<Script>` et son ID de ressources correspondante spécifie l’emplacement du fichier JavaScript avec vos fonctions.
- L'élément `<Page>` et son ID de ressources correspondante spécifie l'emplacement de la page HTML de votre complément. La page HTML comprend un tag `<Script>` qui charge le fichier JavaScript (customfunctions.js). La page HTML est une page masquée qui n’est jamais affichée dans l’interface utilisateur.
- L’élément `<Metadata>` et son ID de ressources correspondante spécifie l’emplacement du fichier JSON.
- Un élément `<Namespace>` et son ID de ressources correspondante spécifie le préfixe pour toutes les fonctions personnalisées dans le complément.


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

## <a name="initializing-custom-functions"></a>Initialisation des fonctions personnalisées

Votre code doit initialiser la fonctionnalité de fonctions personnalisées avant de l'utiliser. Vous pouvez le faire soit dans un tag &lt;Script&gt; dans le fichier HTML (customfunctions.html) ou en haut du fichier JavaScript (customfunctions.js). Lors de l'aperçu des fonctions personnalisées, vous avez le choix entre deux syntaxes pour l'initialisation. Le fichier HTML dans le référentiel utilise la syntaxe suivante :

```js
Office.initialize = function (reason) {
    return Excel.CustomFunctions.initialize();
};
```

Vous pouvez également utiliser la syntaxe suivante :

```js
Office.Preview.StartCustomFunctions();
```

## <a name="handling-errors"></a>Gestion des erreurs
La gestion des erreurs pour les fonctions personnalisées est identique à la [gestion des erreurs pour l’API JavaScript d’Excel dans son ensemble](./excel-add-ins-error-handling.md). En règle générale, vous utiliserez `.catch` pour gérer les erreurs. Le code ci-dessous présente un exemple de `.catch`. 

```js
function getComment(x) {
    var url = "https://jsonplaceholder.typicode.com/comments/" + x; //this delivers a section of lorem ipsum from the jsonplaceholder API
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

## <a name="synchronous-and-asynchronous-functions"></a>Fonctions synchrones et asynchrones

La fonction `ADD42` ci-dessus est synchrone par rapport à Excel (désigné en réglant les paramètres de l'option `"sync": true` dans le fichier JSON). Les fonctions synchrones offrent des performances rapides car elles s’exécutent dans le même processus qu’Excel et s’exécutent en parallèle lors du calcul multithread.   

D'un autre côté, si votre fonction personnalisée récupère des données du Web, elle doit être asynchrone par rapport à Excel. Les fonctions asynchrones doivent :

1. Renvoyer une promesse JavaScript à Excel
3. Résolvez la promesse avec la valeur finale en utilisant la fonction de rappel.

Le code suivant indique un exemple de fonction personnalisée asynchrone qui récupère la température d’un thermomètre. Notez que `sendWebRequest` est une fonction hypothétique, non spécifiée ici, qui utilise XHR pour appeler un service Web de température.

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

Les fonctions asynchrones affichent une erreur temporaire `GETTING_DATA` dans la cellule pendant qu'Excel attend le résultat final. Les utilisateurs peuvent interagir normalement avec le reste du tableur pendant qu’ils attendent le résultat.

> [!NOTE]
> Les fonctions personnalisées sont asynchrones par défaut. Pour désigner les fonctions comme synchrones, définissez l’option `"sync": true` dans la propriété `options` pour la fonction personnalisée dans le fichier JSON d’enregistrement.

## <a name="streamed-functions"></a>Fonctions de flux

Une fonction asynchrone peut être diffusée. Les fonctions personnalisées de flux vous permettent d’afficher des données dans des cellules à plusieurs reprises au fil du temps, sans devoir attendre qu’Excel ou que des utilisateurs demandent à effectuer le calcul à nouveau. L’exemple suivant est une fonction personnalisée qui ajoute un nombre au résultat toutes les secondes. Tenez compte des informations suivantes :

- Excel affiche automatiquement chaque nouvelle valeur en utilisant le rappel `setResult`.
- Le paramètre final, `handler`, n’est jamais spécifié dans votre code d’enregistrement et ne s’affiche pas dans le menu de remplissage automatique pour les utilisateurs d’Excel lorsqu’ils entrent la fonction. Il s’agit d’un objet contenant une fonction de rappel `setResult` utilisée pour transmettre des données de la fonction à Excel afin de mette à jour la valeur d’une cellule.
- Pour qu'Excel transmette la fonction `setResult` dans l'objet `handler`, vous devez déclarer la prise en charge de la diffusion en continu pendant l’enregistrement de votre fonction en définissant l’option `"stream": true` dans la propriété `options` pour la fonction personnalisée dans le fichier JSON d’enregistrement.

```js
function incrementValue(increment, handler){
    var result = 0;
    setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);
}
```

## <a name="cancellation"></a>Annulation

Vous pouvez annuler les fonctions de flux et les fonctions asynchrones. L’annulation de vos appels de fonction permet de considérablement réduire leur consommation de bande passante, la mémoire de travail et la charge de l’UC. Excel annule les appels de fonction dans les situations suivantes :

- L’utilisateur modifie ou supprime une cellule qui fait référence à la fonction.
- Un des arguments (entrées) de la fonction est modifié. Dans ce cas, un nouvel appel de fonction est déclenché en plus de l’annulation.
- L’utilisateur déclenche le nouveau processus de calcul manuellement. Comme pour le cas précédent, un nouvel appel de fonction est déclenché en plus de l’annulation.

Vous *devez* implémenter un gestionnaire d'annulation pour chaque fonction de diffusion en continu. Les fonctions asynchrones, non diffusées en continu peuvent être annulables ou non, c'est à vous de décider. Les fonctions synchrones ne peuvent pas être annulées.

Pour rendre une fonction annulable, définissez l’option `"cancelable": true` dans la propriété `options` pour la fonction personnalisée dans le fichier JSON d’enregistrement.

Le code suivant affiche l’exemple précédent avec l’annulation mise en œuvre. Dans le code, l’objet `handler` contient une fonction `onCanceled` qui doit être définie pour chaque fonction personnalisée pouvant être annulée.

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

## <a name="saving-and-sharing-state"></a>Enregistrement et partage de l'état

Les fonctions asynchrones peuvent enregistrer des données dans des variables JavaScript globales. Lors d’appels ultérieurs, votre fonction personnalisée peut utiliser les valeurs enregistrées dans ces variables. L'état enregistré est utile lorsque les utilisateurs ajoutent la même fonction personnalisée à plusieurs cellules, car toutes les instances de la fonction peuvent partager l'état. Par exemple, vous pouvez enregistrer les données renvoyées par un appel à une ressource web pour éviter d’effectuer des appels supplémentaires à la même ressource web.

Le code suivant illustre une implémentation de la fonction de flux précédente relative à la température qui enregistre l’état global. Tenez compte des informations suivantes :

- `refreshTemperature` est une fonction de flux qui lit la température d’un thermomètre spécifique à chaque seconde qui passe. Les nouvelles températures sont enregistrées dans la variable `savedTemperatures`, mais ne mettent pas directement à jour la valeur de la cellule. Elles ne doivent pas être appelées directement à partir d'une cellule de feuille de calcul, *de sorte qu'elles ne sont pas enregistrées dans le fichier JSON*.
- `streamTemperature` met à jour les valeurs de température affichées dans la cellule chaque seconde et utilise une variable `savedTemperatures` comme source de données. Elles doivent être enregistrées dans le fichier JSON et nommées en lettres majuscules, `STREAMTEMPERATURE`.
- Les utilisateurs peuvent appeler `streamTemperature` à partir de plusieurs cellules dans l’interface utilisateur Excel. Chaque appel lit des données de la même variable `savedTemperatures`.

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

> [!NOTE]
> Les fonctions synchrones (désignées en paramétrant l'option `"sync": true` dans le fichier JSON) ne peuvent pas partager l'état car Excel les parallélise lors du calcul multithread. Seules les fonctions asynchrones peuvent partager l'état car les fonctions synchrones d'un complément partagent le même contexte JavaScript dans chaque session.

## <a name="working-with-ranges-of-data"></a>Utilisation des plages de données

Votre fonction personnalisée accepte les plages de données en tant que paramètres. Sinon, vous pouvez renvoyer une plage de données à partir d’une fonction personnalisée.

Par exemple, supposons que votre fonction renvoie la deuxième valeur la plus élevée prise dans une plage de nombres stockés dans Excel. La fonction suivante prend le paramètre `values`, c’est-à-dire un type de paramètre `Excel.CustomFunctionDimensionality.matrix`. Notez que dans l'enregistrement JSON pour cette fonction, vous devez définir le paramètre propriété `type`sur `matrix`.

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

Comme vous pouvez le voir, les plages sont gérées en JavaScript sous la forme de tableaux de tableaux de lignes (comme un tableau à deux dimensions).

## <a name="known-issues"></a>Problèmes connus

- Les descriptions de paramètre et les URL d’aide ne sont pas encore utilisées par Excel.
- Les fonctions personnalisées ne sont actuellement pas disponibles sur Excel pour les clients mobiles.
- Actuellement, les compléments s’appuient sur un processus de navigateur masqué pour exécuter les fonctions personnalisées asynchrones. À l’avenir, JavaScript s’exécutera directement sur certaines plateformes pour garantir que les fonctions personnalisées sont plus rapides et utilisent moins de mémoire. Par ailleurs, la page HTML référencée par l’élément `<Page>`dans le manifeste ne sera pas nécessaire pour la plupart des plateformes, car Excel exécutera directement le code JavaScript. Pour vous préparer à ce changement, vérifiez que vos fonctions personnalisées n’utilisent pas le DOM de page web. Les API hôtes prises en charge pour accéder au Web seront [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API) et [XHR](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) en utilisant GET ou POST.
- Les fonctions volatiles (celles qui recalculent automatiquement lorsque des modifications de données indépendantes sont effectuées dans le tableur) ne sont pas encore prises en charge.
- Le débogage est uniquement activé pour les fonctions asynchrones sur Excel pour Windows.
- Le déploiement via le portail d'administration Office 365 et AppSource n'est pas encore activé.
- Les fonctions personnalisées dans Excel Online peuvent cesser de fonctionner pendant une session après une période d'inactivité. Actualisez la page du navigateur (F5) et entrez à nouveau une fonction personnalisée pour restaurer la fonction.

## <a name="changelog"></a>Journal des modifications

- **7 novembre 2017 :** mise à disposition des exemples et de la version d’évaluation des fonctions personnalisées
- **20 novembre 2017 :** correction du bogue de compatibilité pour les utilisateurs de la version 8801 et ultérieure
- **28 novembre 2017 :** prise en charge de l’annulation sur des fonctions asynchrones (nécessite la modification des fonctions de flux)
- **7 mai 2018**: Support fourni pour Mac, Excel Online et fonctions synchrones en cours de traitement

\* vers le canal Office Insiders
