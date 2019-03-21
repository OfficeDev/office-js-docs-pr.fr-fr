---
ms.date: 03/19/2019
description: Créer des fonctions personnalisées dans Excel à l’aide de JavaScript.
title: Créer des fonctions personnalisées dans Excel (aperçu)
localization_priority: Priority
ms.openlocfilehash: 4a9e240646b41b737652b6e64eb83e03d0824178
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691201"
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

Par exemple, le code suivant définit les fonctions personnalisées `add` et `increment`indique ensuite les informations de mappage pour les deux fonctions. La fonction`add` mappée à l’objet dans le fichier de métadonnées JSON où la valeur de la `id` propriété est **AJOUTER**et la fonction`increment`mappée à l’objet dans le fichier de métadonnées dans laquelle la valeur de la `id` propriété est **INCRÉMENT**. Voir [Recommandations fonctions personnalisées](custom-functions-best-practices.md#associating-function-names-with-json-metadata) pour plus d’informations sur le mappage des noms de fonction dans le fichier de script pour objets dans le fichier de métadonnées JSON.

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

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
 CustomFunctions.associate("INCREMENT", increment);
```

### <a name="json-metadata-file"></a>Fichier de métadonnées JSON

Le fichier de métadonnées fonctions personnalisées (**./config/customfunctions.json** du projet créé par le Générateur de Yo Office) fournit les informations dont Excel a besoin pour enregistrer les fonctions personnalisées et les rendre disponibles aux utilisateurs finaux. Les fonctions personnalisées sont enregistrées lorsqu’un utilisateur lance un complément pour la première fois. Après cela, elles sont disponibles pour cet utilisateur depuis tous les classeurs (c'est-à-dire pas seulement dans le classeur dans lequel le complément est initialement exécuté.)

> [!TIP]
> Les paramètres du serveur qui héberge le fichier JSON doivent avoir [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) activée afin que les fonctions personnalisées s’exécutent correctement dans Excel Online.

Le code suivant de **customfunctions.json** spécifie les métadonnées pour la `add` fonction et la `increment` fonction qui ont été décrites précédemment. Le tableau qui suit cet exemple de code fournit des informations détaillées sur les propriétés individuelles au sein de cet objet JSON. Voir [Recommandations fonctions personnalisées](custom-functions-best-practices.md#associating-function-names-with-json-metadata) pour plus d’informations sur la spécification de la valeur de `id` et les propriétés`name`dans le fichier de métadonnées JSON.

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
| `options` | Vous permet de personnaliser certains aspects de comment et quand Excel exécute la fonction. Pour plus d’informations sur l’utilisation de cette propriété, consultez les sections [Fonctions de diffusion en continu](#streaming-functions) et [Annulation d’une fonction](#canceling-a-function). |

### <a name="manifest-file"></a>Fichier manifeste

Le fichier manifeste XML pour un complément qui définit les fonctions personnalisées (**./manifest.xml** du projet créé par le Générateur de Yo Office) spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers HTML, JavaScript et JSON. Le balisage XML suivant montre un exemple des éléments`<ExtensionPoint>` et `<Resources>` que vous devez inclure dans manifeste d’un complément pour activer les fonctions personnalisées.  

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

## <a name="declaring-a-volatile-function"></a>Déclaration d’une fonction volatile

Les [fonctions volatiles](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) sont des fonctions dont la valeur change d’un moment à l’autre, même si aucun des arguments de la fonction n’a été modifié. Ces fonctions sont recalculées à chaque recalcul d’Excel. Par exemple, imaginons une cellule qui appelle la fonction `NOW`. Chaque fois que la fonction `NOW` est appelée, elle renvoie automatiquement la date et l’heure actuelles.

Excel contient plusieurs fonctions volatiles intégrées, comme `RAND` et `TODAY`. Pour obtenir la liste complète des fonctions volatiles d’Excel, reportez-vous à [Fonctions volatiles et non volatiles](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

Les fonctions personnalisées permettent de créer vos propres fonctions volatiles, qui peuvent être utiles lors de la gestion des dates, des heures, des nombres aléatoires et de la modélisation. Par exemple, les simulations Monte Carlo exigent la génération d’entrées aléatoires afin de déterminer une solution optimale.

Pour déclarer une fonction volatile, ajoutez `"volatile": true` au sein de l’objet `options` pour la fonction dans le fichier de métadonnées JSON, comme indiqué dans l’exemple de code suivant. Notez qu’une fonction ne peut pas être marquée à la fois `"streaming": true` et `"volatile": true`. Dans le cas où les deux sont marquées comme `true`, l’option volatile est ignorée.

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

## <a name="saving-and-sharing-state"></a>Enregistrement et partage d’état

Les fonctions personnalisées peuvent enregistrer des données dans des variables JavaScript globales, qui peuvent être utilisées dans les appels suivants. Un état enregistré est utile lorsque les utilisateurs appellent la même fonction personnalisée à partir de plusieurs cellules, car toutes les instances de la fonction pouvant accéder à l’état. Par exemple, vous pouvez enregistrer les données renvoyées par un appel à une ressource web pour éviter d’effectuer des appels supplémentaires à la même ressource web.

L’exemple de code suivant montre une implémentation d’une fonction de diffusion en continu de la température qui enregistre l’état global. Tenez compte des informations suivantes à propos de ce code :

- La fonction`streamTemperature`met à jour la valeur de température qui s’affiche dans la cellule chaque seconde et elle utilise la `savedTemperatures` variable en tant que source de données.

- Étant donné que `streamTemperature` est une fonction de diffusion en continu, elle implémente un gestionnaire d’annulation qui s’exécute lorsque la fonction est annulée.

- Si un utilisateur appelle la `streamTemperature` fonction à partir de plusieurs cellules dans Excel, la`streamTemperature` fonction lit les données dans la même `savedTemperatures` variable à chaque fois qu’elle s’exécute. 

- La `refreshTemperature` fonction lit la température d’un thermomètre spécifique à chaque seconde qui passe et stocke le résultat dans la`savedTemperatures`variable. Étant donné que la `refreshTemperature` fonction n’est pas exposée aux utilisateurs finaux dans Excel, elle n’a pas besoin d’être enregistrée dans le fichier JSON.

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

## <a name="coauthoring"></a>Co-création
Excel Online et Excel pour Windows avec un abonnement Office 365 vous permettent de co-créer des documents et cette fonctionnalité est disponible avec les fonctions personnalisées. Si votre classeur utilise une fonction personnalisée, votre collègue sera invité à charger le complément de la fonction personnalisée. Quand vous avez tous les deux chargé le complément, la fonction personnalisée peut partager les résultats via la co-création.

Pour plus d’informations sur la co-création, voir [À propos de la co-création dans Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).

## <a name="working-with-ranges-of-data"></a>Utilisation des plages de données

Votre fonction personnalisée peut accepter une plage de données sous la forme d’un paramètre d’entrée, ou il peut renvoyer une plage de données. Dans JavaScript, une plage de données est représentée sous la forme d’une matrice à deux dimensions.

Par exemple, supposons que votre fonction renvoie la seconde valeur la plus élevée à partir d’une plage de nombres stockés dans Excel. La fonction suivante prend le paramètre `values`, c’est-à-dire un type de `Excel.CustomFunctionDimensionality.matrix`. Notez que dans les métadonnées JSON pour cette fonction, vous devez définir la propriété `type` de paramètre sur `matrix`.

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

## <a name="determine-which-cell-invoked-your-custom-function"></a>Déterminer quelle cellule a appelé votre fonction personnalisée.

Dans certains cas, vous devez récupérer l’adresse de la cellule qui a appelé votre fonction personnalisée. Cela peut être utile dans les types de scénarios suivants:

- Mise en forme de plages: utilisez comme clé la cellule pour stocker des informations dans[AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data). Utilisez ensuite [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) dans Excel pour charger la clé à partir de l’élément `AsyncStorage`.
- Affichage de valeurs mises en cache : si votre fonction est utilisée en mode hors connexion, affichez les valeurs mises en cache à partir de l’élément `AsyncStorage` à l’aide de `onCalculated`.
- Rapprochement : utilisez l’adresse de la cellule pour découvrir la cellule d’origine afin de vous aider à réaliser un rapprochement lors du traitement.

Les informations relatives à l’adresse d’une cellule sont exposées uniquement si `requiresAddress` est marqué comme `true` dans le fichier de métadonnées JSON de la fonction. L’exemple de code suivant illustre ce concept :

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

Dans le fichier de script (**./src/customfunctions.js** ou **./src/customfunctions.ts**), vous devrez également ajouter une fonction `getAddress` pour trouver l’adresse d’une cellule. Cette fonction peut utiliser des paramètres, comme illustré dans l’exemple suivant en tant que `parameter1`. Le dernier paramètre sera toujours `invocationContext`, un objet contenant l’emplacement de la cellule qu’Excel transmet lorsque `requiresAddress` est marqué comme `true` dans votre fichier de métadonnées JSON.

```js
function getAddress(parameter1, invocationContext) {
    return invocationContext.address;
}
```

Par défaut, les valeurs renvoyées par une fonction `getAddress` ont le format suivant : `SheetName!CellNumber`. Par exemple, si une fonction a été appelée à partir d’une feuille de calcul appelée Dépenses dans la cellule B2, la valeur renvoyée serait `Expenses!B2`.

## <a name="known-issues"></a>Problèmes connus

Consulter les problèmes connus sur notre[repo GitHub Fonctions Excel Personnalisées](https://github.com/OfficeDev/Excel-Custom-Functions/issues). 

## <a name="see-also"></a>Voir aussi

* [Métadonnées fonctions personnalisées](custom-functions-json.md)
* [Exécution de fonctions personnalisées Excel](custom-functions-runtime.md)
* [Meilleures pratiques de fonctions personnalisées](custom-functions-best-practices.md)
* [Fonctions personnalisées changelog](custom-functions-changelog.md)
* [Didacticiel de fonctions personnalisées Excel](../tutorials/excel-tutorial-create-custom-functions.md)
