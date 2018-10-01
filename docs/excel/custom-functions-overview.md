---
ms.date: 09/27/2018
description: Créez une fonction personnalisée dans Excel à l’aide de JavaScript.
title: Créer des fonctions personnalisées dans Excel (Aperçu)
ms.openlocfilehash: 98e418f843f6f5574088cea9c7393afc4a42060b
ms.sourcegitcommit: 1852ae367de53deb91d03ca55d16eb69709340d3
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/29/2018
ms.locfileid: "25348800"
---
# <a name="create-custom-functions-in-excel-preview"></a>Créer des fonctions personnalisées dans Excel (aperçu)

Les fonctions personnalisées permettent aux développeurs d'ajouter de nouvelles fonctions à Excel en définissant ces fonctions dans JavaScript comme partie d’un complément. Les utilisateurs d'Excel peuvent accéder à des fonctions personnalisées comme n'importe quelle fonction native d'Excel, telle que `SUM()`. Cet article explique comment créer des fonctions personnalisées dans Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

L’illustration suivante montre un utilisateur final insérant une fonction personnalisée dans une cellule d’une feuille de calcul Excel. La fonction personnalisée `CONTOSO.ADD42` est conçue pour ajouter 42 à la paire de nombres spécifiée par l’utilisateur comme paramètres d’entrée de la fonction.

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

Le code suivant définit la fonction personnalisée `ADD42`.

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> Plus loin dans cet article, la section [Problèmes connus](#known-issues) indique les limites actuelles des fonctions personnalisées.

## <a name="components-of-a-custom-functions-add-in-project"></a>Composants d’un projet de complément de fonctions personnalisées

Si vous utilisez le [générateur de Yo Office](https://github.com/OfficeDev/generator-office) pour créer un projet de complément de fonctions personnalisées Excel, vous verrez les fichiers suivants dans le projet que le générateur crée :

| Fichier | Format de fichier | Description |
|------|-------------|-------------|
| **./src/customfunctions.js**<br/>ou<br/>**./src/customfunctions.ts** | JavaScript<br/>ou<br/>TypeScript | Contient le code qui définit les fonctions personnalisées. |
| **./config/customfunctions.json** | JSON | Contient des métadonnées qui décrivent les fonctions personnalisées et permettent à Excel d'enregistrer les fonctions personnalisées et de les mettre à la disposition des utilisateurs finaux. |
| **./index.html** | HTML | Fournit une référence de &lt;script&gt; pour le fichier JavaScript qui définit les fonctions personnalisées. |
| **./manifest.xml** | XML | Spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers JavaScript, JSON et HTML répertoriés précédemment dans ce tableau. |

Les sections suivantes fournissent plus d’informations sur ces fichiers.

### <a name="script-file"></a>Fichier de script 

Le fichier de script (**./src/customfunctions.js** ou **./src/customfunctions.ts** dans le projet que le générateur de Yo Office crée) contient le code qui définit les fonctions personnalisées et mappe les noms des fonctions personnalisées aux objets du [fichier de métadonnées JSON](#json-metadata-file). 

Par exemple, le code suivant définit les fonctions personnalisées `add` et `increment`, puis spécifie les informations de mappage pour les deux fonctions. La fonction `add` est mappée à l'objet dans le fichier de métadonnées JSON où la valeur de la propriété `id` est **ADD**, et la fonction `increment` est mappée à l'objet dans le fichier de métadonnées où la valeur de la propriété `id` est **INCREMENT**. Pour plus d’informations sur le mappage des noms de fonction dans le fichier de script aux objets dans le fichier de métadonnées JSON, reportez-vous à la rubrique [Meilleures pratiques des fonctions personnalisées](custom-functions-best-practices.md#mapping-function-names-to-json-metadata).

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

Le fichier de métadonnées des fonctions personnalisées (**./config/customfunctions.json** dans le projet que le générateur de Yo Office crée) fournit les informations dont Excel a besoin pour enregistrer les fonctions personnalisées et les rendre disponibles aux utilisateurs finaux. Les fonctions personnalisées sont enregistrées lorsqu’un utilisateur exécute un complément pour la première fois. Après cela, elles sont disponibles pour cet utilisateur dans tous les classeurs (autrement dit, pas seulement dans le classeur dans lequel le complément a été exécuté pour la première fois.)

> [!TIP]
> Parmi les paramètres de serveur sur le serveur qui héberge le fichier JSON, [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) doit être activé pour que les fonctions personnalisées fonctionnent correctement dans Excel Online.

Le code suivant dans **customfunctions.json** spécifie les métadonnées de la fonction `add` et de la fonction `increment` décrites précédemment. Le tableau qui suit cet échantillon de code fournit des informations détaillées sur les propriétés individuelles dans cet objet JSON. Reportez-vous à la rubrique [Meilleures pratiques des fonctions personnalisées](custom-functions-best-practices.md#mapping-function-names-to-json-metadata), pour plus d'informations sur la spécification de la valeur des propriétés `id` et `name` dans le fichier de métadonnées JSON.

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

Le tableau suivant répertorie les propriétés qui sont généralement présentes dans le fichier de métadonnées JSON. Pour plus d'informations détaillées sur le fichier de métadonnées JSON, voir [Métadonnées des fonctions personnalisées](custom-functions-json.md).

| Propriété  | Description |
|---------|---------|
| `id` | ID unique de la fonction. Cet ID ne doit pas être modifié après sa définition. |
| `name` | Nom de la fonction que l’utilisateur final voit dans Excel. Dans Excel, ce nom de fonction sera préfixé par l'espace de noms des fonctions personnalisées, spécifié dans le [fichier manifeste XML](#manifest-file). |
| `helpUrl` | URL de la page qui s’affiche lorsqu’un utilisateur demande de l’aide. |
| `description` | Décrit ce que fait la fonction. Cette valeur s’affiche comme une info-bulle lorsque la fonction est l’élément sélectionné dans le menu de saisie semi-automatique dans Excel. |
| `result`  | Objet qui définit le type de l’information renvoyée par la fonction. La valeur de la propriété enfant `type` peut être **string**, **number**ou **boolean**. La valeur de la propriété enfant `dimensionality` peut être **scalaire** ou **matrice** (un tableau à deux dimensions des valeurs de `type` spécifié). |
| `parameters` | Tableau qui définit les paramètres d’entrée de la fonction. Les propriétés enfants `name` et `description` apparaissent dans l’intelliSense Excel. La valeur de la propriété enfant`type` peut être une **chaîne**, un **nombre**, ou une valeur **booléenne**. La valeur de la propriété enfant `dimensionality` peut être **scalaire** ou **matrice** (un tableau à deux dimensions des valeurs de `type` spécifié). |
| `options` | Vous permet de personnaliser certains aspects de la façon dont Excel exécute la fonction et quand. Pour plus d’informations sur l’utilisation de cette propriété, voir [Fonctions de flux](#streamed-functions) et [Annulation d'une fonction](#canceling-a-function), plus loin dans cet article. |

### <a name="manifest-file"></a>Fichier manifeste

Le fichier manifeste XML pour un complément qui définit les fonctions personnalisées (**./manifest.xml** dans le projet que le générateur de Yo Office crée) spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers JavaScript, JSON et HTML. La balise XML suivante montre un exemple des éléments `<ExtensionPoint>` et `<Resources>` que vous devez inclure dans le manifeste d'un complément pour activer les fonctions personnalisées.  

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
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. -->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> Les fonctions dans Excel sont précédées par l’espace de noms spécifié dans votre fichier manifeste XML. L’espace de noms d’une fonction précède le nom de la fonction, et ils sont séparés par un point. Par exemple, pour appeler la fonction `ADD42` dans la cellule d’une feuille de calcul Excel, vous devez taper `=CONTOSO.ADD42`, puisque CONTOSO est l’espace de noms et `ADD42` est le nom de la fonction spécifiée dans le fichier JSON. L’espace de noms est destiné à être utilisé comme identificateur pour votre entreprise ou le complément. 

## <a name="functions-that-return-data-from-external-sources"></a>Fonctions qui renvoient des données provenant de sources externes

Si une fonction personnalisée récupère les données d’une source externe comme le Web, elle doit :

1. Renvoyer une promesse JavaScript à Excel.

2. Résoudre la promesse avec la valeur finale en utilisant la fonction de rappel.

Les fonctions personnalisées affichent un résultat temporaire `#GETTING_DATA` dans la cellule pendant qu’Excel attend le résultat final. Les utilisateurs peuvent interagir normalement avec le reste de la feuille de calcul tout en attendant le résultat.

Dans l’échantillon de code suivant, la fonction personnalisée `getTemperature()` récupère la température actuelle d’un thermomètre. Remarquez que `sendWebRequest` est une fonction hypothétique (non spécifiée ici) qui utilise [XHR](custom-functions-runtime.md#xhr) pour appeler un service web de température.

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a>Fonctions de flux

Les fonctions de flux personnalisées vous permettent de transmettre des données aux cellules de manière répétée au fil du temps, sans qu'un utilisateur ait à demander explicitement une actualisation des données. L’échantillon de code suivant est une fonction personnalisée qui ajoute un nombre au résultat toutes les secondes. Tenez compte des informations suivantes relatives à ce code :

- Excel affiche automatiquement chaque nouvelle valeur en utilisant le rappel `setResult`.

- Le second paramètre d’entrée, `handler`, n’est pas affiché pour les utilisateurs finaux dans Excel lorsqu’ils sélectionnent la fonction à partir du menu de saisie semi-automatique.

- Le rappel `onCanceled` définit la fonction qui s’exécute lorsque la fonction est annulée. Vous devez implémenter un gestionnaire d'annulation comme celui-ci pour toute fonction de flux. Pour plus d’informations, voir [Annulation d’une fonction](#canceling-a-function). 

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

Lorsque vous spécifiez des métadonnées pour une fonction de flux dans le fichier de métadonnées JSON, vous devez définir les propriétés `"cancelable": true` et `"stream": true` dans l'objet `options`, comme illustré dans l’exemple suivant.

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

Dans certains cas, vous devrez peut-être annuler l’exécution d’une fonction personnalisée en flux continu pour réduire la consommation de la bande passante, de la mémoire et de la charge processeur. Excel annule l’exécution d’une fonction dans les situations suivantes :

- Quand l’utilisateur modifie ou supprime une cellule qui fait référence à la fonction.

- Quand un des arguments (entrées) de la fonction est modifié. Dans ce cas, un nouvel appel de fonction est déclenché suite à l'annulation.

- Lorsque l’utilisateur déclenche manuellement un nouveau calcul. Dans ce cas, un nouvel appel de fonction est déclenché suite à l'annulation.

Pour activer la possibilité d’annuler une fonction, vous devez implémenter un gestionnaire d’annulation dans la fonction JavaScript et spécifier la propriété `"cancelable": true`   dans l'objet `options`   dans les métadonnées JSON qui décrit la fonction. Les échantillons de code dans la section précédente de cet article fournissent un exemple de ces techniques.

## <a name="saving-and-sharing-state"></a>Enregistrement et partage de l'état

Les fonctions personnalisées peuvent enregistrer des données dans des variables JavaScript globales. Lors d’appels ultérieurs, votre fonction personnalisée pourra utiliser les valeurs enregistrées dans ces variables. L'état enregistré est utile lorsque les utilisateurs ajoutent la même fonction personnalisée à plusieurs cellules, car toutes les instances de la fonction peuvent partager l'état. Par exemple, vous pouvez enregistrer les données renvoyées par un appel à une ressource web pour éviter de passer des appels supplémentaires à la même ressource web.

L'échantillon de code suivant montre l'implémentation d'une fonction de flux de température qui enregistre l'état de manière globale. Tenez compte des informations suivantes relatives à ce code :

- `refreshTemperature` ,est une fonction de flux qui chaque seconde, lit la température d’un thermomètre spécifique. Les nouvelles températures sont enregistrées dans la variable `savedTemperatures`, mais ne mettent pas directement à jour la valeur de la cellule. Elles ne doivent pas être appelées directement à partir d'une cellule de feuille de calcul, *de sorte qu'elles ne sont pas enregistrées dans le fichier JSON*.

- `streamTemperature` met à jour les valeurs de température affichées dans la cellule, chaque seconde, et utilise une variable `savedTemperatures` comme source de données. Elles doivent être enregistrées dans le fichier JSON et nommées en lettres majuscules, `STREAMTEMPERATURE`.

- Les utilisateurs peuvent appeler `streamTemperature` à partir de plusieurs cellules dans l’interface utilisateur Excel. Chaque appel lit des données depuis la même variable `savedTemperatures`.

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

## <a name="working-with-ranges-of-data"></a>Utilisation des plages de données

Votre fonction personnalisée peut accepter une plage de données comme paramètre d’entrée, ou elle peut renvoyer une plage de données. En JavaScript, une plage de données est représentée sous la forme d’un tableau à deux dimensions.

Par exemple, supposons que votre fonction renvoie la deuxième valeur la plus élevée prise dans une plage de nombres stockés dans Excel. La fonction suivante accepte le paramètre `values`, qui est de type `Excel.CustomFunctionDimensionality.matrix`. Notez que dans les métadonnées JSON de cette fonction, vous devez définir la propriété `type` du paramètre à `matrix`.

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

Lorsque vous créez un complément qui définit des fonctions personnalisées, veillez à inclure la logique de gestion des erreurs pour prendre en compte les erreurs d’exécution. La gestion des erreurs pour les fonctions personnalisées est identique à la [gestion des erreurs pour l’API JavaScript d’Excel dans son ensemble](excel-add-ins-error-handling.md). Dans l’exemple de code suivant, `.catch` gère les erreurs qui se produisent précédemment dans le code.

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

- Les descriptions de paramètre et les URL d’aide ne sont pas encore utilisées par Excel.
- Les fonctions personnalisées ne sont actuellement pas disponibles sur Excel pour les clients mobiles.
- Les fonctions volatiles (celles qui recalculent automatiquement lorsque des modifications de données indépendantes sont effectuées dans la feuille de calcul) ne sont pas encore prises en charge.
- Le déploiement via le portail d'administration Office 365 et AppSource n'est pas encore activé.
- Les fonctions personnalisées dans Excel Online peuvent cesser de fonctionner pendant une session après une période d'inactivité. Actualisez la page du navigateur (F5) et entrez à nouveau une fonction personnalisée pour restaurer la fonction.
- Il est possible d’avoir le résultat temporaire **#GETTING_DATA** dans la ou les cellules d’une feuille de calcul si vous avez plusieurs compléments s’exécutant dans Microsoft Excel pour Windows. Fermez toutes les fenêtres Excel et redémarrez Excel.
- Des outils de débogage spécifiques pour les fonctions personnalisées pourraient devenir disponibles à l’avenir. En attendant, vous pouvez déboguer sur Excel Online à l’aide des outils de développement F12. Voir plus de détails dans [Meilleures pratiques pour les fonctions personnalisées](custom-functions-best-practices.md).

## <a name="changelog"></a>Journal des modifications

- **7 novembre 2017 :** mise à disposition* de la préversion des fonctions personnalisées et d'exemples
- **20 novembre 2017**: correction du bogue de compatibilité pour les utilisateurs de la version 8801 et ultérieure
- **28 novembre 2017 :** mise à disposition* de la prise en charge de l’annulation sur des fonctions asynchrones (nécessite la modification des fonctions de flux)
- ** 7 mai 2018 : support* fourni pour Mac, Excel Online et fonctions synchrones en cours de **traitement
- **20 septembre 2018** : Support fourni pour les fonctions personnalisées à l'exécution de JavaScript. Pour plus d’informations, voir [Exécution des fonctions personnalisées d’Excel](custom-functions-runtime.md).

\* vers le canal Office Insiders

## <a name="see-also"></a>Voir aussi

* [Métadonnées des fonctions personnalisées](custom-functions-json.md)
* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
* [Meilleures pratiques pour les fonctions personnalisées](custom-functions-best-practices.md)
* [Didacticiel sur les fonctions personnalisées d’Excel](excel-tutorial-custom-functions.md)