---
ms.date: 03/29/2019
description: Créer des fonctions personnalisées dans Excel à l’aide de JavaScript.
title: Créer des fonctions personnalisées dans Excel (aperçu)
localization_priority: Priority
ms.openlocfilehash: 7a461728061ace532a11a8473d27ec4340eebb97
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448471"
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

Si vous utilisez le [générateur de Yo Office](https://github.com/OfficeDev/generator-office) pour créer un projet de complément de fonctions personnalisées Excel, vous constaterez qu’il crée des fichiers qui contrôlent totalement vos fonctions, votre volet des tâches et votre complément. Nous allons vous concentrer sur les fichiers importants pour les fonctions personnalisées : 

| File | Format de fichier | Description |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>ou<br/>**./src/functions/functions.ts** | JavaScript<br/>ou<br/>TypeScript | Contient le code qui définit les fonctions personnalisées. |
| **./src/functions/functions.html** | HTML | Fournit une référence&lt;script&gt; au fichier JavaScript qui définit les fonctions personnalisées. |
| **./manifest.xml** | XML | Spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers JavaScript et HTML qui figurent plus haut dans ce tableau. Répertorie également les emplacements des autres fichiers que votre complément pourrait utiliser, tels que les fichiers du volet des tâches et les fichiers de commande. |

### <a name="script-file"></a>Fichier de script

Le fichier de script (**./src/functions/functions.js** ou **./src/functions/functions.ts** dans le projet que crée le générateur de Yo Office) contient le code qui définit des fonctions personnalisées, des commentaires qui définissent la fonction, et associe les noms des fonctions personnalisées à des objets dans le fichier de métadonnées JSON.

Le code suivant définit la fonction personnalisée `add`, puis spécifie des informations d’association pour la fonction. Pour plus d’informations sur l’association de fonctions, voir [Meilleures pratiques des fonctions personnalisées](custom-functions-best-practices.md#associating-function-names-with-json-metadata).

Le code suivant fournit également des commentaires de code qui définissent la fonction. Le commentaire obligatoire `@customfunction` est déclaré en premier, pour indiquer qu’il s’agit d’une fonction personnalisée. Vous pouvez également constater que deux paramètres sont déclarés, `first` et `second`, qui sont suivis de leurs propriétés `description`. Enfin, une description `returns` est fournie. Pour plus d’informations sur les commentaires requis pour votre fonction personnalisée, voir [Générer des métadonnées JSON pour des fonctions personnalisées](custom-functions-json-autogeneration.md).

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
```

### <a name="manifest-file"></a>Fichier manifeste

Le fichier manifeste XML pour un complément qui définit les fonctions personnalisées (**./manifest.xml** du projet créé par le Générateur de Yo Office) spécifie l’espace de noms pour toutes les fonctions personnalisées dans le complément et l’emplacement des fichiers HTML, JavaScript et JSON. 

Le marquage XML suivant présente un exemple des éléments`<ExtensionPoint>` et `<Resources>` que vous devez inclure dans le manifeste d’un complément pour activer les fonctions personnalisées. Si vous utilisez le générateur de Yo Office, vos fichiers de fonction personnalisée générés contiennent un fichier manifeste plus complexe que vous pouvez comparer sur [ce dépôt Github](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml).

> [!NOTE] 
> Les URL spécifiées dans le fichier manifeste pour les fonctions personnalisées de fichiers HTML, JavaScript et JSON doivent avoir le même sous-domaine et être accessibles publiquement.

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
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
        <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
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

Dans le fichier de script (**./src/functions/functions.js** ou **./src/functions/functions.ts**), vous devrez également ajouter une fonction `getAddress` pour trouver l’adresse d’une cellule. Cette fonction peut utiliser des paramètres, comme illustré dans l’exemple suivant en tant que `parameter1`. Le dernier paramètre sera toujours `invocationContext`, un objet contenant l’emplacement de la cellule qu’Excel transmet lorsque `requiresAddress` est marqué comme `true` dans votre fichier de métadonnées JSON.

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
* [Débogage des fonctions personnalisées](custom-functions-debugging.md)
