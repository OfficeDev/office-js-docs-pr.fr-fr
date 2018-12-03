---
ms.date: 10/17/2018
description: Créer des fonctions personnalisées dans Excel à l’aide de JavaScript.
title: Créer des fonctions personnalisées dans Excel (Aperçu)
ms.openlocfilehash: 8383b5f6d568a1ce2da036fbacfb90404bbe8297
ms.sourcegitcommit: 2ac7d64bb2db75ace516a604866850fce5cb2174
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/14/2018
ms.locfileid: "26298550"
---
# <a name="create-custom-functions-in-excel-preview"></a>Créer des fonctions personnalisées dans Excel (aperçu)

Les fonctions personnalisées permettent aux développeurs d’ajouter de nouvelles fonctions dans Excel en définissant ces fonctions dans JavaScript dans le cadre d’un complément. Les utilisateurs d’Excel peuvent accéder aux fonctions personnalisées comme ils le feraient pour n’importe quelle fonction native d’Excel, telle que `SUM()`. Cet article explique comment créer des fonctions personnalisées dans Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

L’illustration suivante montre un utilisateur final insérant une fonction personnalisée dans une cellule de feuille de calcul Excel. Le `CONTOSO.ADD42` fonction personnalisée est conçue pour ajouter 42 à la paire de nombres que spécifie l’utilisateur en tant que paramètres d’entrée de la fonction.

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

Le code suivant définit la `ADD42` fonction personnalisée.

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> La section [problèmes connus](#known-issues)plus loin dans cet article indique les limitations en cours de fonctions personnalisées.

## <a name="components-of-a-custom-functions-add-in-project"></a>Composants d’un projet de complément fonctions personnalisées

Si vous utilisez le [générateur Yo Office](https://github.com/OfficeDev/generator-office) pour créer un projet complément de fonctions personnalisées Excel, vous verrez les fichiers suivants dans le projet crée par le générateur :

| Fichier | Format de fichier | Description |
|------|-------------|-------------|
| **./src/customfunctions.js**<br/>ou<br/>**./src/customfunctions.ts** | JavaScript<br/>ou<br/>TypeScript | Contient le code qui définit les fonctions personnalisées. |
| **./config/customfunctions.json** | JSON | Contient les métadonnées qui décrivent les fonctions personnalisées et permettent à Excel d’enregistrer les fonctions personnalisées et les rendre accessibles aux utilisateurs finaux. |
| **./index.html** | HTML | Fournit une référence&lt;script&gt; au fichier JavaScript qui définit les fonctions personnalisées. |
| **./manifest.xml** | XML | Spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers HTML, JavaScript et JSON qui figurent précédemment dans ce tableau. |

Les sections suivantes vous apportent plus d'informations sur ces fichiers.

### <a name="script-file"></a>Fichier de script 

Le fichier de script (**./src/customfunctions.js** ou **./src/customfunctions.ts** du projet créé par le Générateur de Yo Office) contient le code qui définit les fonctions personnalisées et mappe les noms des fonctions personnalisées aux objets dans le [fichier de métadonnées JSON](#json-metadata-file). 

Par exemple, le code suivant définit les fonctions personnalisées `add` et `increment`indique ensuite les informations de mappage pour les deux fonctions. La fonction `add` mappée à l’objet dans le fichier de métadonnées JSON où la valeur de la `id` propriété est **AJOUTER**et la fonction`increment`mappée à l’objet dans le fichier de métadonnées dans laquelle la valeur de la `id` propriété est **INCRÉMENT**. Voir [Recommandations fonctions personnalisées](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) pour plus d’informations sur le mappage des noms de fonction dans le fichier de script pour objets dans le fichier de métadonnées JSON.

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

### <a name="json-metadata-file"></a>Fichier de métadonnées JSON 

Le fichier de métadonnées fonctions personnalisées (**./config/customfunctions.json** du projet créé par le Générateur de Yo Office) fournit les informations dont Excel a besoin pour enregistrer les fonctions personnalisées et les rendre disponibles aux utilisateurs finaux. Les fonctions personnalisées sont enregistrées lorsqu’un utilisateur lance un complément pour la première fois. Après cela, elles sont disponibles pour cet utilisateur depuis tous les classeurs (c'est-à-dire pas seulement dans le classeur dans lequel le complément est initialement exécuté.)

> [!TIP]
> Les paramètres du serveur qui héberge le fichier JSON doivent avoir [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) activée afin que les fonctions personnalisées s’exécutent correctement dans Excel Online.

Le code suivant de **customfunctions.json** spécifie les métadonnées pour la `add` fonction et la `increment` fonction qui ont été décrites précédemment. Le tableau qui suit cet exemple de code fournit des informations détaillées sur les propriétés individuelles au sein de cet objet JSON. Voir [Recommandations fonctions personnalisées](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) pour plus d’informations sur la spécification de la valeur de `id` et les propriétés`name`dans le fichier de métadonnées JSON.

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

Le tableau suivant répertorie les propriétés généralement présentes dans le fichier de métadonnées JSON. Pour plus d’informations sur le fichier de métadonnées JSON, voir [métadonnées fonctions personnalisées](custom-functions-json.md).

| Propriété  | Description |
|---------|---------|
| `id` | Un ID unique pour la fonction. Cet ID peut contenir uniquement des points et caractères alphanumériques et ne doit pas être modifié une fois défini. |
| `name` | Nom de la fonction que voit l’utilisateur final dans Excel. Dans Excel, ce nom de la fonction sera précédé de l’espace de noms de fonctions personnalisées spécifié dans le [fichier manifeste XML](#manifest-file). |
| `helpUrl` | URL de la page qui s’affiche quand un utilisateur demande de l’aide. |
| `description` | Descriptif de la fonction. Cette valeur apparaît comme une info-bulle lorsque la fonction est l’élément sélectionné dans le menu de saisie semi-automatique des formules dans Excel. |
| `result`  | Objet qui définit le type d’informations renvoyées par la fonction. Pour plus d’informations sur cet objet, voir [résultat](custom-functions-json.md#result). |
| `parameters` | Tableau qui définit les paramètres d’entrée de la fonction. Pour plus d’informations sur cet objet, voir [paramètres](custom-functions-json.md#parameters). |
| `options` | Vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction. Pour plus d’informations sur l’utilisation de cette propriété, voir [Diffusion en continu de fonctions](#streaming-functions) et [Annuler une fonction](#canceling-a-function) plus loin dans cet article. |

### <a name="manifest-file"></a>Fichier manifeste

Le fichier manifeste XML pour un complément qui définit les fonctions personnalisées (**./manifest.xml** du projet créé par le Générateur de Yo Office) spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers HTML, JavaScript et JSON. Le balisage XML suivant montre un exemple des éléments`<ExtensionPoint>` et `<Resources>` que vous devez inclure dans manifeste d’un complément pour activer les fonctions personnalisées.  

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
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. Can only contain alphanumeric characters and periods.-->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> Les fonctions dans Excel sont précédées par l’espace de noms spécifié dans votre fichier manifeste XML. L’espace de noms d’une fonction vient avant le nom de fonction et les deux sont séparés par un point. Par exemple, pour appeler la fonction `ADD42` dans la cellule de feuille de calcul Excel, vous saisiriez `=CONTOSO.ADD42`, car `CONTOSO` est l’espace de noms et `ADD42` est le nom de la fonction spécifié dans le fichier JSON. L’espace de noms est destiné à être utilisé comme identificateur de votre entreprise ou du complément. Un espace de noms ne peut contenir que des points et des caractères alphanumériques.

## <a name="functions-that-return-data-from-external-sources"></a>Fonctions qui retournent des données provenant de sources externes

Si une fonction personnalisée récupère des données d’une source externe comme le web, elle doit :

1. Renvoyer une promesse JavaScript à Excel.

2. Résoudre la promesse avec la valeur finale à l’aide de la fonction de rappel.

Les fonctions personnalisées affichent un `#GETTING_DATA`résultat temporaire dans la cellule, tandis qu’ Excel attend que le résultat final. Les utilisateurs peuvent interagir normalement avec le reste de la feuille de calcul pendant qu’ils attendent le résultat.

Le code suivant indique un exemple de`getTemperature()`fonction personnalisée qui récupère la température d’un thermomètre. Notez que `sendWebRequest` est une fonction hypothétique (non spécifiée ici) qui utilise [XHR](custom-functions-runtime.md#xhr-example) pour appeler un service web de température.

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streaming-functions"></a>Fonctions de diffusion en continu

Les fonctions personnalisées de diffusion en continu vous aident à copier des données à des cellules à plusieurs reprises au fil du temps, sans exiger qu’un utilisateur demande explicitement l’actualisation des données. L’exemple de code suivant est une fonction personnalisée qui ajoute un nombre au résultat chaque seconde. Tenez compte des informations suivantes à propos de ce code :

- Excel affiche chaque nouvelle valeur automatiquement à l’aide du `setResult` rappel.

- Le deuxième paramètre d’entrée `handler`, n’est pas visible aux utilisateurs finaux dans Excel lorsqu’ils sélectionnent la fonction à partir du menu de saisie semi-automatique.

- Le `onCanceled` rappel définit la fonction qui s’exécute lorsque la fonction est annulée. Vous devez implémenter un gestionnaire d’annulation comme suit pour n’importe quelle fonction de diffusion en continu. Pour plus d’informations, voir [Annuler une fonction](#canceling-a-function).

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

Lorsque vous spécifiez des métadonnées pour une fonction de diffusion en continu dans le fichier de métadonnées JSON, vous devez définir les propriétés `"cancelable": true` et `"stream": true` au sein de l’objet`options`, comme illustré dans l’exemple suivant.

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

## <a name="canceling-a-function"></a>Annulation d’une fonction

Dans certains cas, vous devrez annuler l’exécution d’une fonction personnalisée de diffusion en continu pour réduire la consommation de bande passante, de la mémoire de travail et la charge du CPU. Excel annule l’exécution d’une fonction dans les situations suivantes :

- L’utilisateur modifie ou supprime une cellule qui fait référence à la fonction.

- Un des arguments (entrées) de la fonction est modifié. Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.

- L’utilisateur déclenche manuellement le recalcul. Dans ce cas, un appel de nouvelle fonction est déclenché en plus de l’annulation.

Pour activer la possibilité d’annuler une fonction, vous devez implémenter un gestionnaire d’annulation au sein de la fonction JavaScript et spécifier la propriété `"cancelable": true` au sein de l’objet`options` dans les métadonnées JSON décrivant la fonction. Les exemples de code dans la section précédente de cet article fournissent un exemple de ces techniques.

## <a name="saving-and-sharing-state"></a>Enregistrement et partage d’état

Les fonctions personnalisées peuvent enregistrer des données dans des variables JavaScript globales, qui peuvent être utilisées dans les appels suivants. Un état enregistré est utile lorsque les utilisateurs appellent la même fonction personnalisée à partir de plusieurs cellules, car toutes les instances de la fonction pouvant accéder à l’état. Par exemple, vous pouvez enregistrer les données renvoyées par un appel à une ressource web pour éviter d’effectuer des appels supplémentaires à la même ressource web.

L’exemple de code suivant montre une implémentation d’une fonction de diffusion en continu de la température qui enregistre l’état global. Tenez compte des informations suivantes à propos de ce code :

- La fonction`streamTemperature`met à jour la valeur de température qui s’affiche dans la cellule chaque seconde et elle utilise la `savedTemperatures` variable en tant que source de données.

- Étant donné que `streamTemperature` est une fonction de diffusion en continu, elle implémente un gestionnaire d’annulation qui s’exécute lorsque la fonction est annulée.

- Si un utilisateur appelle la `streamTemperature` fonction à partir de plusieurs cellules dans Excel, la`streamTemperature` fonction lit les données dans la même `savedTemperatures` variable à chaque fois qu’elle s’exécute. 

- La `refreshTemperature` fonction lit la température d’un thermomètre spécifique à chaque seconde qui passe et stocke le résultat dans la`savedTemperatures`variable. Étant donné que la `refreshTemperature` fonction n’est pas exposée aux utilisateurs finaux dans Excel, elle ne doit pas être enregistrées dans le fichier JSON.

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

## <a name="working-with-ranges-of-data"></a>Utilisation des plages de données

Votre fonction personnalisée peut accepter une plage de données sous la forme d’un paramètre d’entrée, ou il peut renvoyer une plage de données. Dans JavaScript, une plage de données est représentée sous la forme d’une matrice 2 dimensions.

Par exemple, supposons que votre fonction renvoie la seconde valeur la plus élevée à partir d’une plage de nombres stockés dans Excel. La fonction suivante prend le paramètre `values`, c’est-à-dire un type de `Excel.CustomFunctionDimensionality.matrix`. Notez que dans les métadonnées JSON pour cette fonction, vous devez définir la propriété `type` de paramètre à `matrix`.

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

## <a name="handling-errors"></a>Gestion des erreurs

Lorsque vous créez un complément à l’aide des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution. La gestion des erreurs pour fonctions personnalisées est identique à la[gestion des erreurs pour l’API JavaScript Excel](excel-add-ins-error-handling.md). Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.

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

## <a name="known-issues"></a>Problèmes connus

- Les descriptions de paramètre et les URL d’aide ne sont pas encore utilisés par Excel.
- Les fonctions personnalisées ne sont actuellement pas disponibles dans Excel pour les clients mobiles.
- Les fonctions volatiles (celles qui sont recalculées à chaque fois que des données autonomes sont modifiées dans la feuille de calcul) ne sont pas encore prises en charge.
- Le déploiement via le portail d’administration Office 365 et AppSource n’est pas encore activé.
- Les fonctions personnalisées dans Excel Online peuvent cesser de fonctionner pendant une session après une période d’inactivité. Actualiser la page du navigateur (F5), puis entrez une fonction personnalisée pour restaurer la fonctionnalité.
- Vous pouvez voir le **## CHARGEMENT_DONNEES** résultat temporaire au sein des cellules d’une feuille de calcul si vous avez plusieurs compléments en cours d’exécution sur Excel pour Windows. Fermez toutes les fenêtres Excel et redémarrez Excel.
- Des outils de débogage spécifiques aux fonctions personnalisées seront peut-être disponibles à l’avenir. En attendant, vous pouvez déboguer sur Excel Online à l’aide des outils de développement F12. Plus de détails dans [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md).

## <a name="changelog"></a>Journal des modifications

- **7 novembre 2017 :** mise à disposition des exemples et de l’aperçu des fonctions personnalisées
- **20 novembre 2017 :** correction du bogue de compatibilité pour les utilisateurs de la version 8801 et ultérieure
- **28 novembre 2017 :** prise en charge de l’annulation sur des fonctions asynchrones (nécessite la modification des fonctions de flux)
- **7 mai 2018**: prise en charge pour Mac, Excel Online et fonctions synchrones dans les processus en cours d’exécution
- **20 septembre 2018**: prise en charge de fonctions personnalisées lors de l’exécution de JavaScript. Pour plus d’informations, voir [Runtime pour les fonctions personnalisées Excel](custom-functions-runtime.md).
- **20 octobre 2018**: avec le programme[October Insiders build](https://support.office.com/fr-FR/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24), les fonctions personnalisées nécessitent désormais le paramètre « id » dans votre[métadonnées fonctions personnalisées](custom-functions-json.md) pour les versions Windows Bureau et Online. Sur Mac, ce paramètre doit être ignoré.


\* pour la chaîne [Office Insider](https://products.office.com/office-insider) (anciennement appelée « Insider Fast »)

## <a name="see-also"></a>Voir aussi

* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
* [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md)
* [Didacticiel de fonctions personnalisées Excel](excel-tutorial-custom-functions.md)
