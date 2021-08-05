---
title: Programmation asynchrone dans des compléments Office
description: Découvrez comment la bibliothèque JavaScript Office utilise la programmation asynchrone dans Office’applications.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 6408d1efc99f38468b371247156d84f1a4ac4b99
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773943"
---
# <a name="asynchronous-programming-in-office-add-ins"></a>Programmation asynchrone dans des compléments Office

[!include[information about the common API](../includes/alert-common-api-info.md)]

Pourquoi l’API de Compléments Office a-t-elle recours à la programmation asynchrone ? JavaScript étant un langage monothread, si le script appelle un processus synchrone de longue durée, toute exécution de script ultérieure sera bloquée tant que ce processus ne sera pas terminé. Étant donné que certaines opérations sur les clients web Office (mais également sur les clients riches) peuvent bloquer l’exécution si elles sont exécutées de manière synchrone, la plupart des API JavaScript Office sont conçues pour s’exécuter de manière asynchrone. Cela permet de s’assurer Office les modules sont réactifs et rapides. Vous devez donc fréquemment écrire des fonctions de rappel lorsque vous utilisez ces méthodes asynchrones.

Les noms de toutes les méthodes asynchrones dans l’API se terminent par « Async », par `Document.getSelectedDataAsync` `Binding.getDataAsync` exemple, ou `Item.loadCustomPropertiesAsync` les méthodes. Lorsqu’une méthode « Async » est appelée, elle est exécutée immédiatement et toute exécution de script ultérieure peut se poursuivre normalement. La fonction de rappel facultative que vous transmettez à une méthode « Async » s’exécute dès que l’opération demandée ou les données sont prêtes. L’opération est généralement rapide, mais le retour pourrait présenter un léger retard.

Le diagramme suivant illustre le flux d’exécution d’un appel à une méthode « Async » qui lit les données sélectionnées par l’utilisateur dans un document ouvert dans le serveur Word ou Excel. Au moment où l’appel « Async » est effectué, le thread d’exécution JavaScript est libre d’effectuer tout traitement côté client supplémentaire (bien qu’aucun ne soit affiché dans le diagramme). Lorsque la méthode « Async » est de retour, le rappel reprend l’exécution sur le thread et le module peut accéder aux données, y faire quelque chose et afficher le résultat. Le même modèle d’exécution asynchrone est valable lorsque vous travaillez avec Office applications clientes enrichies, telles que Word 2013 ou Excel 2013.

*Figure 1. Flux d’exécution de programmation asynchrone*

![Diagramme montrant l’interaction d’exécution de commande au fil du temps avec l’utilisateur, la page de la application et le serveur d’applications web hébergeant le module.](../images/office-addins-asynchronous-programming-flow.png)

La prise en charge de cette conception asynchrone dans les clients riches et les clients web fait partie des objectifs de conception « écriture unique-exécution multiplateforme » du modèle de développement des Compléments Office. Par exemple, vous pouvez créer un complément de contenu ou du volet de tâches avec une seule base de code qui sera exécutée sur Excel 2013 et Excel sur le web.

## <a name="write-the-callback-function-for-an-async-method"></a>Écrire la fonction de rappel pour une méthode « Async »

La fonction de rappel que  vous passez en tant qu’argument de rappel à une méthode « Async » doit déclarer un seul paramètre que le runtime du add-in utilisera pour fournir l’accès à un objet [AsyncResult](/javascript/api/office/office.asyncresult) lors de l’exécution de la fonction de rappel. Vous pouvez écrire:

- Fonction anonyme qui doit être écrite et transmise directement en ligne avec l’appel à la méthode « Async » comme paramètre de rappel de la méthode « Async ». 

- Fonction nommée, en passant le nom  de cette fonction comme paramètre de rappel d’une méthode « Async ».

Une fonction anonyme est utile si vous envisagez de n’utiliser son code qu’une fois : comme elle n’a pas de nom, vous ne pouvez pas y faire référence dans une autre partie du code. Une fonction nommée est utile si vous voulez réutiliser la fonction de rappel pour plusieurs méthodes « Async ».

### <a name="write-an-anonymous-callback-function"></a>Écrire une fonction de rappel anonyme

La fonction de rappel anonyme suivante déclare un seul paramètre nommé qui récupère les données de la propriété `result` [AsyncResult.value](/javascript/api/office/office.asyncresult#value) lors du retour du rappel.

```js
function (result) {
        write('Selected data: ' + result.value);
}
```

L’exemple suivant montre comment transmettre cette fonction de rappel anonyme en ligne dans le contexte d’un appel complet de méthode « Async » à la `Document.getSelectedDataAsync` méthode.

- Le premier argument _coercionType,_ spécifie de renvoyer les données sélectionnées sous forme `Office.CoercionType.Text` de chaîne de texte.

- Le deuxième argument _de rappel_ est la fonction anonyme transmise en ligne à la méthode. Lorsque la fonction s’exécute, elle utilise le paramètre de résultat pour accéder à la propriété de l’objet afin d’afficher les données sélectionnées par  `value` l’utilisateur dans le `AsyncResult` document.

```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
    function (result) {
        write('Selected data: ' + result.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Vous pouvez également utiliser le paramètre de votre fonction de rappel pour accéder à d’autres propriétés de `AsyncResult` l’objet. Utilisez la propriété [AsyncResult.status](/javascript/api/office/office.asyncresult#status) pour déterminer si l’appel a réussi ou échoué. En cas d’échec, vous pouvez utiliser la propriété [AsyncResult.error](/javascript/api/office/office.asyncresult#error) pour accéder à un objet [Error](/javascript/api/office/office.error) et obtenir des informations sur l’erreur.

Pour plus d’informations sur l’utilisation de la méthode, voir Lire et écrire des données dans la sélection active dans `getSelectedDataAsync` un document ou une feuille de [calcul.](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md) 

### <a name="write-a-named-callback-function"></a>Écrire une fonction de rappel nommée

Vous pouvez également écrire une fonction nommée et  passer son nom au paramètre de rappel d’une méthode « Async ». Par exemple, l’exemple précédent peut être réécrit pour passer une fonction nommée `writeDataCallback` en tant que paramètre _callback_ comme suit.

```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
    writeDataCallback);

// Callback to write the selected data to the add-in UI.
function writeDataCallback(result) {
    write('Selected data: ' + result.value);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```


## <a name="differences-in-whats-returned-to-the-asyncresultvalue-property"></a>Différences dans les éléments retournés à la propriété AsyncResult.value

Les propriétés et les propriétés de l’objet retournent les mêmes types d’informations à la fonction de rappel transmise à toutes les méthodes `asyncContext` `status` « `error` `AsyncResult` Async ». Toutefois, ce qui est renvoyé à la propriété varie en fonction des fonctionnalités de la méthode `AsyncResult.value` « Async ».

Par exemple, les méthodes (des objets `addHandlerAsync` [Binding](/javascript/api/office/office.binding), [CustomXmlPart,](/javascript/api/office/office.customxmlpart) [Document,](/javascript/api/office/office.document) [RoamingSettings](/javascript/api/outlook/office.roamingsettings)et [Paramètres)](/javascript/api/office/office.settings) sont utilisées pour ajouter des fonctions de handler d’événements aux éléments représentés par ces objets. Vous pouvez accéder à la propriété à partir de la fonction de rappel que vous passez à l’une des méthodes, mais comme aucune donnée ou objet n’est accessible lorsque vous ajoutez un handler d’événements, la propriété renvoie toujours `AsyncResult.value` `addHandlerAsync` `value` **undefined** si vous tentez d’y accéder.

En revanche, si vous appelez la méthode, elle renvoie les données sélectionnées par l’utilisateur dans le document à la propriété `Document.getSelectedDataAsync` `AsyncResult.value` dans le rappel. Ou, si vous appelez la méthode [Bindings.getAllAsync,](/javascript/api/office/office.bindings#getAllAsync_options__callback_) elle renvoie un tableau de tous les objets `Binding` du document. Si vous appelez la méthode [Bindings.getByIdAsync,](/javascript/api/office/office.bindings#getByIdAsync_id__options__callback_) elle renvoie un seul `Binding` objet.

Pour obtenir une description de ce qui est renvoyé à la propriété pour une méthode, voir la section « Valeur de rappel » de la rubrique de référence `AsyncResult.value` `Async` de cette méthode. Pour obtenir un résumé de tous les objets qui fournissent des méthodes, consultez le tableau en bas de la rubrique sur `Async` l’objet [AsyncResult.](/javascript/api/office/office.asyncresult)

## <a name="asynchronous-programming-patterns"></a>Modèles de programmation asynchrone

L Office API JavaScript prend en charge deux types de modèles de programmation asynchrone.

- Utilisation des rappels imbriqués
- Utilisation du modèle des promesses

La programmation asynchrone à l’aide des fonctions de rappel nécessite que vous imbriquiez fréquemment le résultat retourné d’un rappel au sein d’au moins deux rappels. Pour ce faire, vous pouvez utiliser les rappels imbriqués de toutes les méthodes « Async » de l’API.

L’utilisation des rappels imbriqués est un modèle de programmation familier pour la plupart des développeurs JavaScript, mais le code contenant des rappels fortement imbriqués peut être difficile à lire et à comprendre. Comme alternative aux rappels imbrmbrés, l’API JavaScript Office prend également en charge une implémentation du modèle de promesses.

> [!NOTE]
> Dans la version actuelle de l’API JavaScript *Office,* la prise en charge intégrée du modèle de promesses fonctionne uniquement avec du code pour les liaisons dans les feuilles de calcul Excel et les [documents Word.](bind-to-regions-in-a-document-or-spreadsheet.md) Toutefois, vous pouvez encapsuler d’autres fonctions qui ont des rappels à l’intérieur de votre propre fonction de renvoi de promesse personnalisée. Pour plus d’informations, voir [Wrap Common APIs in Promise-returning functions](#wrap-common-apis-in-promise-returning-functions).

### <a name="asynchronous-programming-using-nested-callback-functions"></a>Programmation asynchrone utilisant des fonctions de rappel imbriquées

Vous devez fréquemment effectuer au moins deux opérations asynchrones pour réaliser une tâche. Pour ce faire, vous pouvez imbriquer un appel « Async » dans un autre.

L’exemple de code suivant imbrique deux appels asynchrones.

- D’abord, la méthode [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getByIdAsync_id__options__callback_) est appelée pour accéder à une liaison dans le document nommé « MyBinding ». L’objet renvoyé au paramètre de ce rappel permet d’accéder à l’objet de `AsyncResult` liaison spécifié à partir de la `result` `AsyncResult.value` propriété.
- Ensuite, l’objet de liaison accessible à partir du premier paramètre est utilisé pour appeler la méthode `result` [Binding.getDataAsync.](/javascript/api/office/office.binding#getDataAsync_options__callback_)
- Enfin, le paramètre du rappel passé à la méthode est utilisé pour afficher `result2` les données dans la `Binding.getDataAsync` liaison.

```js
function readData() {
    Office.context.document.bindings.getByIdAsync("MyBinding", function (result) {
        result.value.getDataAsync({ coercionType: 'text' }, function (result2) {
            write(result2.value);
        });
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Ce modèle de rappel imbrique de base peut être utilisé pour toutes les méthodes asynchrones dans l Office API JavaScript.

Les sections suivantes montrent comment utiliser des fonctions anonymes ou nommées pour des rappels imbriqués dans des méthodes asynchrones.

#### <a name="use-anonymous-functions-for-nested-callbacks"></a>Utiliser des fonctions anonymes pour les rappels imbrmbrés

Dans l’exemple suivant, deux fonctions anonymes sont déclarées en ligne et transmises dans les méthodes et dans les `getByIdAsync` `getDataAsync` rappels imbrmbrés. Comme les fonctions sont très simples, l’objet de l’implémentation est évident.

```js
Office.context.document.bindings.getByIdAsync('myBinding', function (bindingResult) {
    bindingResult.value.getDataAsync(function (getResult) {
        if (getResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        } else {
            write('Data has been read successfully.');
        }
    });
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

#### <a name="use-named-functions-for-nested-callbacks"></a>Utiliser des fonctions nommées pour les rappels imbrmbrés

Dans des implémentations complexes, il peut être utile d’utiliser des fonctions nommées pour garantir une meilleure lisibilité, simplicité de gestion et possibilité de réutilisation du code. Dans l’exemple suivant, les deux fonctions anonymes de l’exemple de la section précédente ont été réécrites en tant que fonctions nommées `deleteAllData` et `showResult` . Ces fonctions nommées sont ensuite passées dans les méthodes et dans les `getByIdAsync` `deleteAllDataValuesAsync` rappels par nom.

```js
Office.context.document.bindings.getByIdAsync('myBinding', deleteAllData);

function deleteAllData(asyncResult) {
    asyncResult.value.deleteAllDataValuesAsync(showResult);
}

function showResult(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Data has been deleted successfully.');
    }
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

### <a name="asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings"></a>Programmation asynchrone en utilisant le modèle des promesses pour accéder aux données des liaisons

Plutôt que de transmettre une fonction de rappel et d’attendre le renvoi de la fonction pour poursuivre l’exécution, le motif de programmation des promesses renvoie immédiatement un objet de promesse qui représente le résultat souhaité. Toutefois, contrairement à la vraie programmation synchrone, en arrière-plan, la concrétisation du résultat prévu est en fait différée jusqu’à ce que l’environnement d’exécution des compléments Office puisse réaliser la demande. Un gestionnaire _onError_ est fourni pour couvrir les cas où la demande ne peut pas être remplie.

L Office API JavaScript fournit la [méthode Office.select](/javascript/api/office#Office_select_expression__callback_) pour prendre en charge le modèle de promesses permettant d’utiliser des objets de liaison existants. L’objet promise renvoyé à la méthode prend en charge uniquement les quatre méthodes accessibles directement à partir de l’objet Binding : `Office.select` [getDataAsync](/javascript/api/office/office.binding#getDataAsync_options__callback_), [setDataAsync](/javascript/api/office/office.binding#setDataAsync_data__options__callback_), [addHandlerAsync](/javascript/api/office/office.binding#addHandlerAsync_eventType__handler__options__callback_)et [removeHandlerAsync](/javascript/api/office/office.binding#removeHandlerAsync_eventType__options__callback_). [](/javascript/api/office/office.binding)

Le modèle de promesses pour l’travail avec les liaisons prend cette forme.

**Office.select(**_selectorExpression_, _onError_**).** _BindingObjectAsyncMethod_

Le paramètre _selectorExpression_ prend la forme , où bindingId est le nom ( ) d’une liaison que vous avez créée précédemment dans le document ou la feuille de calcul (à l’aide de l’une des méthodes `"bindings#bindingId"` «  `id` addFrom » de la collection : `Bindings` , ou `addFromNamedItemAsync` `addFromPromptAsync` `addFromSelectionAsync` ). Par exemple, l’expression du sélecteur spécifie que vous souhaitez accéder à la liaison avec `bindings#cities` **l’ID** « cities ».

Le _paramètre onError_ est une fonction de gestion des erreurs qui prend un seul paramètre de type qui peut être utilisé pour accéder à un objet, si la méthode ne parvient pas à accéder à la `AsyncResult` liaison `Error` `select` spécifiée. L’exemple suivant montre une fonction de gestion des erreurs de base pouvant être passée au paramètre _onError_.

```js
function onError(result){
    var err = result.error;
    write(err.name + ": " + err.message);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Remplacez l’espace réservé _BindingObjectAsyncMethod_ par un appel à l’une des quatre méthodes d’objet pris en charge par l’objet `Binding` promise : , , ou `getDataAsync` `setDataAsync` `addHandlerAsync` `removeHandlerAsync` . Les appels à ces méthodes ne prennent pas en charge les promesses supplémentaires. Vous devez les appeler à l’aide du [modèle de fonction de rappel imbriquée](#asynchronous-programming-using-nested-callback-functions).

Une fois qu’une promesse d’objet est remplie, elle peut être réutilisée dans l’appel de méthode chaînée comme s’il s’agit d’une liaison (le runtime du add-in ne retentera pas de manière `Binding` asynchrone la promesse). Si la promesse d’objet ne peut pas être remplie, le runtime du add-in essaiera à nouveau d’accéder à l’objet de liaison la prochaine fois qu’une de ses méthodes `Binding` asynchrones sera invoquée.

L’exemple de code suivant utilise la méthode pour récupérer une liaison avec le « » à partir de la collection, puis appelle la méthode `select` `id` `cities` `Bindings` [addHandlerAsync](/javascript/api/office/office.binding#addHandlerAsync_eventType__handler__options__callback_) [](/javascript/api/office/office.bindingdatachangedeventargs) pour ajouter un handler d’événements pour l’événement dataChanged de la liaison.

```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```

> [!IMPORTANT]
> La `Binding` promesse d’objet renvoyée par la méthode permet d’accéder uniquement `Office.select` aux quatre méthodes de `Binding` l’objet. Si vous devez accéder à l’un des autres membres de l’objet, vous devez utiliser la propriété et ou les méthodes `Binding` `Document.bindings` pour récupérer `Bindings.getByIdAsync` `Bindings.getAllAsync` `Binding` l’objet. Par exemple, si vous devez accéder à l’une des propriétés de l’objet (la ou les propriétés), ou si vous devez accéder aux propriétés des objets `Binding` `document` `id` `type` [MatrixBinding](/javascript/api/office/office.matrixbinding) ou [TableBinding,](/javascript/api/office/office.tablebinding) `getByIdAsync` `getAllAsync` vous devez utiliser la ou les méthodes pour récupérer un `Binding` objet.

## <a name="pass-optional-parameters-to-asynchronous-methods"></a>Passer des paramètres facultatifs à des méthodes asynchrones

La syntaxe courante pour toutes les méthodes « Async » suit ce modèle.

 _AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_`);`

Toutes les méthodes asynchrones prennent en charge des paramètres facultatifs, qui sont passés sous la forme d’un objet JSON (JavaScript Object Notation) qui contient un ou plusieurs paramètres facultatifs. L’objet JSON contenant les paramètres facultatifs est une collection non ordonnée de paires clé-valeur où le caractère « : » sépare la clé de la valeur. Chaque paire dans l’objet est séparée par une virgule, et l’ensemble complet de paires est placé entre accolades. La clé est le nom du paramètre, et la valeur est la valeur à passer pour ce paramètre.

Vous pouvez créer l’objet JSON qui contient des paramètres facultatifs en ligne, ou en créant un objet et en le passant en tant que `options` _paramètre d’options._

### <a name="pass-optional-parameters-inline"></a>Transmettre les paramètres facultatifs en ligne

Par exemple, la syntaxe pour appeler la méthode [Document.setSelectedDataAsync](/javascript/api/office/office.document#setSelectedDataAsync_data__options__callback_) avec des paramètres facultatifs incorporés se présente comme ceci :

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

Dans cette forme de syntaxe d’appel, les deux paramètres _facultatifs, coercionType_ et _asyncContext_, sont définis comme un objet JSON inclus entre accolades.

L’exemple suivant montre comment appeler la `Document.setSelectedDataAsync` méthode en spécifiant des paramètres facultatifs en ligne.

```js
Office.context.document.setSelectedDataAsync(
    "<html><body>hello world</body></html>",
    {coercionType: "html", asyncContext: 42},
    function(asyncResult) {
        write(asyncResult.status + " " + asyncResult.asyncContext);
    }
)

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

> [!NOTE]
> Vous pouvez spécifier des paramètres facultatifs dans l’objet JSON dans n’importe quel ordre dans la mesure où leurs noms sont correctement spécifiés.

### <a name="pass-optional-parameters-in-an-options-object"></a>Passer des paramètres facultatifs dans un objet Options

Vous pouvez également créer un objet nommé qui spécifie les paramètres facultatifs séparément de l’appel de méthode, puis passer l’objet en tant `options` `options` qu’argument _options._

L’exemple suivant montre une façon de créer l’objet, où , et ainsi de suite, sont des espaces réservé pour les noms et valeurs des paramètres `options` `parameter1` `value1` réels.

```js
var options = {
    parameter1: value1,
    parameter2: value2,
    ...
    parameterN: valueN
};

```

Ce qui ressemble à l’exemple suivant lors de la spécification des paramètres [ValueFormat](/javascript/api/office/office.valueformat) et [FilterType](/javascript/api/office/office.filtertype).

```js
var options = {
    valueFormat: "unformatted",
    filterType: "all"
};
```

Voici une autre façon de créer `options` l’objet.

```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

Ce qui ressemble à l’exemple suivant lorsqu’il est utilisé pour spécifier `ValueFormat` les `FilterType` paramètres :

```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```

> [!NOTE]
> Lorsque vous utilisez l’une ou l’autre méthode de création de l’objet, vous pouvez spécifier des paramètres facultatifs dans n’importe quel ordre tant que leurs noms `options` sont spécifiés correctement.

L’exemple suivant montre comment appeler la méthode `Document.setSelectedDataAsync` en spécifiant des paramètres facultatifs dans un `options` objet.

```js
var options = {
   coercionType: "html",
   asyncContext: 42
};

document.setSelectedDataAsync(
    "<html><body>hello world</body></html>",
    options,
    function(asyncResult) {
        write(asyncResult.status + " " + asyncResult.asyncContext);
    }
)

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Dans les deux exemples de paramètres facultatifs, le paramètre de rappel est spécifié en tant que dernier paramètre (en suivant les paramètres facultatifs inline ou en suivant l’objet _d’argument options)._  Vous pouvez également spécifier le paramètre _callback_ à l’intérieur de l’objet JSON incorporé, ou dans l’objet `options`. Cependant, vous ne pouvez passer le paramètre _callback_ qu’à un seul endroit : soit dans l’objet _options_ (incorporé ou créé en externe), soit comme dernier paramètre, mais pas les deux.

## <a name="wrap-common-apis-in-promise-returning-functions"></a>Wrap Common APIs in Promise-returning functions

Les méthodes d’API communes (Outlook API) ne retournent pas [de promesses.](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) Par conséquent, vous ne pouvez pas utiliser [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) pour suspendre l’exécution jusqu’à ce que l’opération asynchrone se termine. Si vous avez `await` besoin d’un comportement, vous pouvez encapsuler l’appel de méthode dans une promesse créée explicitement. 

Le modèle de base consiste à créer une méthode asynchrone qui renvoie un objet Promise  immédiatement et résout cet objet Promise lorsque la méthode interne est terminée ou rejette l’objet en cas d’échec de la méthode.  Voici un exemple simple.

```javascript
function getDocumentFilePath() {
    return new OfficeExtension.Promise(function (resolve, reject) {
        try {
            Office.context.document.getFilePropertiesAsync(function (asyncResult) {
                resolve(asyncResult.value.url);
            });
        }
        catch (error) {
            reject(WordMarkdownConversion.errorHandler(error));
        }
    })
}
```

Lorsque cette méthode doit être attendue, elle peut être appelée avec le mot clé ou en tant que `await` fonction transmise à une `then` fonction.

> [!NOTE]
> Cette technique est particulièrement utile lorsque vous devez appeler l’une des API communes à l’intérieur d’un appel de la méthode dans l’un des modèles objet `run` propres à l’application. Pour obtenir un exemple de la fonction ci-dessus utilisée de cette façon, voir le fichierHome.js dans l’exemple [ Word-Add-in-JavaScript-MDConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion/blob/master/Word-Add-in-JavaScript-MDConversionWeb/Home.js).

Voici un exemple utilisant TypeScript.

```typescript
readDocumentFileAsync(): Promise<any> {
    return new Promise((resolve, reject) => {
        const chunkSize = 65536;
        const self = this;

        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: chunkSize }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(asyncResult.error);
            } else {
                // `getAllSlices` is a Promise-wrapped implementation of File.getSliceAsync.
                self.getAllSlices(asyncResult.value).then(result => {
                    if (result.IsSuccess) {
                        resolve(result.Data);
                    } else {
                        reject(asyncResult.error);
                    }
                });
            }
        });
    });
}
```

## <a name="see-also"></a>Voir aussi

- [Compréhension de l’API JavaScript pour Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript pour Office](../reference/javascript-api-for-office.md)
