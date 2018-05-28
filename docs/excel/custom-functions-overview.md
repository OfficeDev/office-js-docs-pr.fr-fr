# <a name="create-custom-functions-in-excel-preview"></a>Cr?er des fonctions personnalis?es dans Excel (Aper?u)

Les fonctions personnalis?es (similaires aux fonctions d?finies par l?utilisateur) permettent aux d?veloppeurs d?ajouter n?importe quelle fonction JavaScript ? Excel en utilisant un compl?ment. Les utilisateurs peuvent alors avoir acc?s aux fonctions personnalis?es comme toute autre fonction native dans Excel (telle que `=SUM()`). Cet article explique comment cr?er des fonctions personnalis?es dans Excel.

L'illustration suivante montre comment un utilisateur final ins?re une fonction personnalis?e dans une cellule. La fonction qui ajoute 42 ? une paire de nombres.

<img alt="custom functions" src="../images/custom-function.gif" width="579" height="383" />

Voici le code pour la m?me fonction personnalis?e.

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

Les fonctions personnalis?es sont d?sormais disponibles dans Developer Preview sous Windows, Mac et Excel Online. Pour les tester, proc?dez comme suit :

1.  Installez Office (version 9325 sur Windows ou 13.329 sur Mac) et participez au programme [Office Insider](https://products.office.com/en-us/office-insider). (Notez qu'il ne suffit pas d'obtenir la derni?re version, la fonctionnalit? sera d?sactiv?e sur n'importe quelle version jusqu'? ce que vous rejoignez le programme Insider)
2.  Clonez le d?p?t des [fonctions Excel personnalis?es](https://github.com/OfficeDev/Excel-Custom-Functions) et suivez les instructions dans le fichier README.md pour d?marrer le compl?ment dans Excel, apporter des modifications dans le code et d?boguer.
3.  Saisissez `=CONTOSO.ADD42(1,2)` dans une cellule, puis appuyez sur **Entr?e** pour ex?cuter la fonction personnalis?e.

Reportez-vous ? la section **Probl?mes connus**? la fin de cet article qui inclut les limites actuelles des fonctions personnalis?es et sera mise ? jour au fil du temps.

## <a name="learn-the-basics"></a>Notions fondamentales

Dans le d?p?t d?exemple clon?, vous trouverez les fichiers suivants?:

- **customfunctions.js**, qui contient le code de fonction personnalis? (voir l'exemple de code simple ci-dessus pour la fonction `ADD42`).
- **customfunctions.json**, qui contient l?enregistrement JSON qui indique ? Excel votre fonction personnalis?e. Avec l?enregistrement, vos fonctions personnalis?es apparaissent dans la liste des fonctions disponibles affich?e lorsqu'un utilisateur saisit du texte dans les cellules.
- **customfunctions.html**, qui fournit une r?f?rence &lt;Scipt&gt; au fichier JS. Ce fichier n?affiche pas d?interface utilisateur dans Excel.
- **customfunctions.xml**, qui indique ? Excel l?emplacement des fichiers HTML, JavaScript et JSON, et sp?cifie ?galement un espace de noms pour toutes les fonctions personnalis?es install?es avec le compl?ment.

### <a name="json-file-customfunctionsjson"></a>Fichier JSON (customfunctions.json)

Le code suivant dans customfunctions.json sp?cifie les m?tadonn?es pour la m?me fonction `ADD42`.

> [!NOTE]
> Les informations de r?f?rence d?taill?es pour le fichier JSON, y compris les options non utilis?es dans cet exemple, sont dans [Enregistrement des fonctions personnalis?es JSON](https://dev.office.com/reference/add-ins/custom-functions-json).

Notez que pour cet exemple?:

- Il n'y a qu'une seule fonction personnalis?e, donc il n'y a qu'un seul membre d tableau `functions`.
- La propri?t? `name` d?finit le nom de la fonction. Comme vous le voyez dans l'image gif anim?e montr?e pr?c?demment, un espace de noms (`CONTOSO`) est ajout? au nom de la fonction dans le menu remplissage automatique Excel. Ce pr?fixe est d?fini dans le manifeste du compl?ment, d?crit ci-dessous. Le pr?fixe et le nom de la fonction sont s?par?s ? l'aide d'un point et, par convention, les pr?fixes et les noms de fonctions sont en majuscules. Pour utiliser votre fonction personnalis?e, un utilisateur tape l?espace de nom suivi du nom de la fonction (`ADD42`) dans une cellule, dans ce cas `=CONTOSO.ADD42`. Le pr?fixe est destin? ? ?tre utilis? comme identificateur de votre entreprise ou du compl?ment. 
- Le `description` appara?t dans le menu remplissage automatique dans Excel.
- Lorsque l?utilisateur demande de l?aide concernant une fonction, Excel ouvre un volet Office et affiche la page web accessible via cette URL sp?cifi?e dans `helpUrl`.
- La propri?t? `result` sp?cifie le type d?information retourn?e ? Excel par la fonction. La propri?t? enfant `type` peut `"string"`, `"number"`, ou `"boolean"`. La propri?t? `dimensionality` peut ?tre `scalar` ou `matrix` (un tableau bidimensionnel de valeurs de la valeur sp?cifi?e `type`).
- Le tableau `parameters` sp?cifie, *dans l'ordre*, le type de donn?es dans chaque param?tre qui est pass? ? la fonction. Les propri?t?s enfants `name` et `description` sont utilis?es dans l?intelliSense Excel. Les propri?t?s enfants `type` et `dimensionality` sont identiques aux propri?t?s enfants de la propri?t? `result` d?crite ci-dessus.
- La propri?t? `options` vous permet de personnaliser certains aspects de la fa?on dont Excel ex?cute la fonction et quand. Il y a plus d'informations sur ces options plus loin dans cet article.

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
> Les fonctions personnalis?es sont enregistr?es lorsqu?un utilisateur ex?cute le compl?ment pour la premi?re fois. Apr?s cela, elles sont disponibles, pour le m?me utilisateur, dans tous les classeurs (pas seulement celui dans lequel le compl?ment a fonctionn? initialement.)

Vos param?tres de serveur pour le fichier JSON doivent avoir activ? [CORS](https://developer.mozilla.org/en-US/docs/Web/HTTP/CORS) pour que les fonctions personnalis?es fonctionnent correctement dans Excel Online.


### <a name="manifest-file-customfunctionsxml"></a>Fichier manifeste (customfunctions.xml)


Ce qui suit est un exemple de balisage `<ExtensionPoint>` et `<Resources>` ? inclure dans le manifeste du compl?ment pour permettre ? Excel d?ex?cuter vos fonctions. Notez ce qui suit ? propos de ce balisage :

- L??l?ment `<Script>` et son ID de ressources correspondante sp?cifie l?emplacement du fichier JavaScript avec vos fonctions.
- L'?l?ment `<Page>` et son ID de ressources correspondante sp?cifie l'emplacement de la page HTML de votre compl?ment. La page HTML comprend un tag `<Script>` qui charge le fichier JavaScript (customfunctions.js). La page HTML est une page masqu?e qui n?est jamais affich?e dans l?interface utilisateur.
- L??l?ment `<Metadata>` et son ID de ressources correspondante sp?cifie l?emplacement du fichier JSON.
- Un ?l?ment `<Namespace>` et son ID de ressources correspondante sp?cifie le pr?fixe pour toutes les fonctions personnalis?es dans le compl?ment.


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

## <a name="initializing-custom-functions"></a>Initialisation des fonctions personnalis?es

Votre code doit initialiser la fonctionnalit? de fonctions personnalis?es avant de l'utiliser. Vous pouvez le faire soit dans un tag &lt;Script&gt; dans le fichier HTML (customfunctions.html) ou en haut du fichier JavaScript (customfunctions.js). Lors de l'aper?u des fonctions personnalis?es, vous avez le choix entre deux syntaxes pour l'initialisation. Le fichier HTML dans le r?f?rentiel utilise la syntaxe suivante?:

```js
Office.initialize = function (reason) {
    return Excel.CustomFunctions.initialize();
};
```

Vous pouvez ?galement utiliser la syntaxe suivante :

```js
Office.Preview.StartCustomFunctions();
```

## <a name="synchronous-and-asynchronous-functions"></a>Fonctions synchrones et asynchrones

La fonction `ADD42` ci-dessus est synchrone par rapport ? Excel (d?sign? en r?glant les param?tres de l'option `"sync": true` dans le fichier JSON). Les fonctions synchrones offrent des performances rapides car elles s?ex?cutent dans le m?me processus qu?Excel et s?ex?cutent en parall?le lors du calcul multithread.   

D'un autre c?t?, si votre fonction personnalis?e r?cup?re des donn?es du Web, elle doit ?tre asynchrone par rapport ? Excel. Les fonctions asynchrones doivent?:

1. Renvoyer une promesse JavaScript ? Excel.
3. R?solvez la promesse avec la valeur finale en utilisant la fonction de rappel.

Le code suivant indique un exemple de fonction personnalis?e asynchrone qui r?cup?re la temp?rature d?un thermom?tre. Notez que `sendWebRequest` est une fonction hypoth?tique, non sp?cifi?e ici, qui utilise XHR pour appeler un service Web de temp?rature.

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

Les fonctions asynchrones affichent une erreur temporaire `GETTING_DATA` dans la cellule pendant qu'Excel attend le r?sultat final. Les utilisateurs peuvent interagir normalement avec le reste du tableur pendant qu?ils attendent le r?sultat.

> [!NOTE]
> Les fonctions personnalis?es sont asynchrones par d?faut. Pour d?signer les fonctions comme synchrones, d?finissez l?option `"sync": true` dans la propri?t? `options` pour la fonction personnalis?e dans le fichier JSON d?enregistrement.

## <a name="streamed-functions"></a>Fonctions de flux

Une fonction asynchrone peut ?tre diffus?e. Les fonctions personnalis?es de flux vous permettent d?afficher des donn?es dans des cellules ? plusieurs reprises au fil du temps, sans devoir attendre qu?Excel ou que des utilisateurs demandent ? effectuer le calcul ? nouveau. L?exemple suivant est une fonction personnalis?e qui ajoute un nombre au r?sultat toutes les secondes. Tenez compte des informations suivantes?:

- Excel affiche automatiquement chaque nouvelle valeur en utilisant le rappel `setResult`.
- Le param?tre final, `caller`, n?est jamais sp?cifi? dans votre code d?enregistrement et ne s?affiche pas dans le menu de remplissage automatique pour les utilisateurs d?Excel lorsqu?ils entrent la fonction. Il s?agit d?un objet contenant une fonction de rappel `setResult` utilis?e pour transmettre des donn?es de la fonction ? Excel afin de mette ? jour la valeur d?une cellule.
- Pour qu'Excel transmette la fonction `setResult` dans l'objet `caller`, vous devez d?clarer la prise en charge de la diffusion en continu pendant l?enregistrement de votre fonction en d?finissant l?option `"stream": true` dans la propri?t? `options` pour la fonction personnalis?e dans le fichier JSON d?enregistrement.

```js
function incrementValue(increment, caller){
    var result = 0;
    setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);
}
```

## <a name="cancellation"></a>Annulation

Vous pouvez annuler les fonctions de flux et les fonctions asynchrones. L?annulation de vos appels de fonction permet de consid?rablement r?duire leur consommation de bande passante, la m?moire de travail et la charge de l?UC. Excel annule les appels de fonction dans les situations suivantes :

- L?utilisateur modifie ou supprime une cellule qui fait r?f?rence ? la fonction.
- Un des arguments (entr?es) de la fonction est modifi?. Dans ce cas, un nouvel appel de fonction est d?clench? en plus de l?annulation.
- L?utilisateur d?clenche le nouveau processus de calcul manuellement. Comme pour le cas pr?c?dent, un nouvel appel de fonction est d?clench? en plus de l?annulation.

Vous *devez* impl?menter un gestionnaire d'annulation pour chaque fonction de diffusion en continu. Les fonctions asynchrones, non diffus?es en continu peuvent ?tre annulables ou non, c'est ? vous de d?cider. Les fonctions synchrones ne peuvent pas ?tre annul?es.

Pour rendre une fonction annulable, d?finissez l?option `"cancelable": true` dans la propri?t? `options` pour la fonction personnalis?e dans le fichier JSON d?enregistrement.

Le code suivant affiche l?exemple pr?c?dent avec l?annulation mise en ?uvre. Dans le code, l?objet `caller` contient une fonction `onCanceled` qui doit ?tre d?finie pour chaque fonction personnalis?e.

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

## <a name="saving-and-sharing-state"></a>Enregistrement et partage de l'?tat

Les fonctions asynchrones peuvent enregistrer des donn?es dans des variables JavaScript globales. Lors d?appels ult?rieurs, votre fonction personnalis?e peut utiliser les valeurs enregistr?es dans ces variables. L'?tat enregistr? est utile lorsque les utilisateurs ajoutent la m?me fonction personnalis?e ? plusieurs cellules, car toutes les instances de la fonction peuvent partager l'?tat. Par exemple, vous pouvez enregistrer les donn?es renvoy?es par un appel ? une ressource web pour ?viter d?effectuer des appels suppl?mentaires ? la m?me ressource web.

Le code suivant illustre une impl?mentation de la fonction de diffusion en continu pr?c?dente relative ? la temp?rature qui enregistre l??tat ? l?aide la variable. Tenez compte des informations suivantes?:

- `refreshTemperature` est une fonction diffus?e en continu qui lit la temp?rature d?un thermom?tre sp?cifique ? chaque seconde qui passe. Les nouvelles temp?ratures sont enregistr?es dans la variable `savedTemperatures`, mais ne mettent pas directement ? jour la valeur de la cellule. Elles ne doivent pas ?tre appel?es directement ? partir d'une cellule de feuille de calcul, *de sorte qu'elles ne sont pas enregistr?es dans le fichier JSON*.
- `streamTemperature` met ? jour les valeurs de temp?rature affich?es dans la cellule chaque seconde et utilise une variable `savedTemperatures` comme source de donn?es. Elles doivent ?tre enregistr?es dans le fichier JSON et nomm?es en lettres majuscules, `STREAMTEMPERATURE`.
- Les utilisateurs peuvent appeler `streamTemperature` ? partir de plusieurs cellules dans l?interface utilisateur Excel. Chaque appel lit des donn?es de la m?me variable `savedTemperatures`.

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
> Les fonctions synchrones (d?sign?es en param?trant l'option `"sync": true` dans le fichier JSON) ne peuvent pas partager l'?tat car Excel les parall?lise lors du calcul multithread. Seules les fonctions asynchrones peuvent partager l'?tat car les fonctions synchrones d'un compl?ment partagent le m?me contexte JavaScript dans chaque session.

## <a name="working-with-ranges-of-data"></a>Utilisation des plages de donn?es

Votre fonction personnalis?e accepte les plages de donn?es en tant que param?tres. Sinon, vous pouvez renvoyer une plage de donn?es ? partir d?une fonction personnalis?e.

Par exemple, supposons que votre fonction renvoie la deuxi?me valeur la plus ?lev?e parmi une plage de nombres stock?e dans Excel. La fonction suivante prend le param?tre `values`, c?est-?-dire un type de param?tre `Excel.CustomFunctionDimensionality.matrix`. Notez que dans l'enregistrement JSON pour cette fonction, vous devez d?finir le param?tre propri?t? `type`sur `matrix`.

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

Comme vous pouvez le voir, les plages sont g?r?es en JavaScript sous la forme de tableaux de tableaux de lignes (comme un tableau ? deux dimensions).

## <a name="known-issues"></a>Probl?mes connus

- Les descriptions de param?tre et les URL d?aide ne sont pas encore utilis?es par Excel.
- Les fonctions personnalis?es ne sont actuellement pas disponibles sur Excel pour les clients mobiles.
- Actuellement, les compl?ments s?appuient sur un processus de navigateur masqu? pour ex?cuter les fonctions asynchrones. ? l?avenir, JavaScript s?ex?cutera directement sur certaines plateformes pour garantir que les fonctions personnalis?es sont plus rapides et utilisent moins de m?moire. Par ailleurs, la page HTML r?f?renc?e par l??l?ment `<Page>`dans le manifeste ne sera pas n?cessaire pour la plupart des plateformes, car Excel ex?cutera directement le code JavaScript. Pour vous pr?parer ? ce changement, v?rifiez que vos fonctions personnalis?es n?utilisent pas le DOM de page web. Les API h?tes prises en charge pour acc?der au Web seront [WebSocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) et [XHR](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) en utilisant GET ou POST.
- Les fonctions volatiles (celles qui recalculent automatiquement lorsque des modifications de donn?es ind?pendantes sont effectu?es dans le tableur) ne sont pas encore prises en charge.
- Le d?bogage est uniquement activ? pour les fonctions asynchrones sur Excel pour Windows.
- Le d?ploiement via le portail d'administration Office 365 et AppSource n'est pas encore activ?.
- Les fonctions personnalis?es dans Excel Online peuvent cesser de fonctionner pendant une session apr?s une p?riode d'inactivit?. Actualisez la page du navigateur (F5) et entrez ? nouveau une fonction personnalis?e pour restaurer la fonction.

## <a name="changelog"></a>Journal des modifications

- **7 novembre 2017 :** mise ? disposition des exemples et de la version d??valuation des fonctions personnalis?es
- **20 novembre 2017 :** correction du bogue de compatibilit? pour les utilisateurs de la version 8801 et ult?rieure
- **28 novembre 2017 :** prise en charge de l?annulation sur des fonctions asynchrones (n?cessite la modification des fonctions de flux)
- **7 mai 2018**: Support fourni pour Mac, Excel Online et fonctions synchrones en cours de traitement
