---
title: Programmation asynchrone dans des compléments Office
description: Découvrez comment la bibliothèque JavaScript Office utilise la programmation asynchrone dans les compléments Office.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 96805ee0f78caedd718642a97828db26f0de7900
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408578"
---
# <a name="asynchronous-programming-in-office-add-ins"></a>Programmation asynchrone dans des compléments Office

[!include[information about the common API](../includes/alert-common-api-info.md)]

Pourquoi l’API de Compléments Office a-t-elle recours à la programmation asynchrone ? JavaScript étant un langage monothread, si le script appelle un processus synchrone de longue durée, toute exécution de script ultérieure sera bloquée tant que ce processus ne sera pas terminé. Étant donné que certaines opérations sur les clients Web Office (mais aussi les clients enrichis) peuvent bloquer l’exécution si elles sont exécutées de manière synchrone, la plupart des API JavaScript d’Office sont conçues pour s’exécuter de manière asynchrone. Cela permet de s’assurer que les compléments Office sont réactifs et rapides. Vous devez donc fréquemment écrire des fonctions de rappel lorsque vous utilisez ces méthodes asynchrones.

Les noms de toutes les méthodes asynchrones de l’API se terminent par « Async », comme les `Document.getSelectedDataAsync` `Binding.getDataAsync` méthodes, ou `Item.loadCustomPropertiesAsync` . Lorsqu’une méthode « Async » est appelée, elle est exécutée immédiatement et toute exécution de script ultérieure peut se poursuivre normalement. La fonction de rappel facultative que vous transmettez à une méthode « Async » s’exécute dès que l’opération demandée ou les données sont prêtes. L’opération est généralement rapide, mais le retour pourrait présenter un léger retard.

Le diagramme suivant illustre le flux d’exécution d’un appel à une méthode « Async » qui lit les données sélectionnées par l’utilisateur dans un document ouvert dans Word ou Excel. Au moment de l’appel « Async », le thread d’exécution JavaScript est libre d’effectuer tout traitement supplémentaire côté client (même si aucun n’est affiché dans le diagramme). Lorsque la méthode « Async » est renvoyée, le rappel reprend l’exécution sur le thread, et le complément peut accéder aux données, effectuer une opération avec ce dernier et afficher le résultat. Le même modèle d’exécution asynchrone est conservé lorsque vous utilisez les applications client riche Office, telles que Word 2013 ou Excel 2013.

*Figure 1. Flux d’exécution de programmation asynchrone*

![Flux d’exécution de thread de programmation asynchrone](../images/office-addins-asynchronous-programming-flow.png)

La prise en charge de cette conception asynchrone dans les clients riches et les clients web fait partie des objectifs de conception « écriture unique-exécution multiplateforme » du modèle de développement des Compléments Office. Par exemple, vous pouvez créer un complément de contenu ou du volet de tâches avec une seule base de code qui sera exécutée sur Excel 2013 et Excel sur le web.

## <a name="writing-the-callback-function-for-an-async-method"></a>Écriture de la fonction de rappel pour une méthode « Async »


La fonction de rappel transmise en tant qu’argument de _rappel_ à une méthode « Async » doit déclarer un paramètre unique que le runtime de complément utilisera pour fournir l’accès à un objet [asyncResult](/javascript/api/office/office.asyncresult) lors de l’exécution de la fonction de rappel. Vous pouvez écrire:


- Une fonction anonyme qui doit être écrite et passée directement en ligne avec l’appel à la méthode « Async » en tant que paramètre _callback_ de la méthode « Async ».

- Une fonction nommée, en passant le nom de cette fonction en tant que paramètre _callback_ d’une méthode « Async ».

Une fonction anonyme est utile si vous envisagez de n’utiliser son code qu’une fois : comme elle n’a pas de nom, vous ne pouvez pas y faire référence dans une autre partie du code. Une fonction nommée est utile si vous voulez réutiliser la fonction de rappel pour plusieurs méthodes « Async ».


### <a name="writing-an-anonymous-callback-function"></a>Écriture d’une fonction de rappel anonyme

La fonction de rappel anonyme suivante déclare un seul paramètre nommé `result` qui récupère les données à partir de la propriété [asyncResult. Value](/javascript/api/office/office.asyncresult#value) lorsque le rappel est retourné.


```js
function (result) {
        write('Selected data: ' + result.value);
}
```

L’exemple suivant montre comment transmettre cette fonction de rappel anonyme en ligne dans le contexte d’un appel complet de méthode « Async » à la `Document.getSelectedDataAsync` méthode.


- Le premier argument _coercionType_ , `Office.CoercionType.Text` , spécifie de renvoyer les données sélectionnées en tant que chaîne de texte.

- Le deuxième argument de _rappel_ est la fonction anonyme passée dans la ligne à la méthode. Lorsque la fonction s’exécute, elle utilise le paramètre _result_ pour accéder à la `value` propriété de l' `AsyncResult` objet afin d’afficher les données sélectionnées par l’utilisateur dans le document.


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

Vous pouvez également utiliser le paramètre de votre fonction de rappel pour accéder à d’autres propriétés de l' `AsyncResult` objet. Utilisez la propriété [AsyncResult.status](/javascript/api/office/office.asyncresult#status) pour déterminer si l’appel a réussi ou échoué. En cas d’échec, vous pouvez utiliser la propriété [AsyncResult.error](/javascript/api/office/office.asyncresult#error) pour accéder à un objet [Error](/javascript/api/office/office.error) et obtenir des informations sur l’erreur.

Pour plus d’informations sur l’utilisation de la `getSelectedDataAsync` méthode, voir [lecture et écriture de données dans la sélection active d’un document ou d’une feuille de calcul](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md). 


### <a name="writing-a-named-callback-function"></a>Écriture d’une fonction de rappel nommée

Vous pouvez également écrire une fonction nommée et transmettre son nom au paramètre _callback_ d’une méthode « Async ». Par exemple, l’exemple précédent peut être réécrit pour passer une fonction nommée `writeDataCallback` en tant que paramètre _callback_ comme suit.


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


Les `asyncContext` `status` Propriétés, et `error` de l' `AsyncResult` objet renvoient les mêmes types d’informations à la fonction de rappel transmise à toutes les méthodes « Async ». Toutefois, ce qui est renvoyé à la `AsyncResult.value` propriété varie en fonction de la fonctionnalité de la méthode « Async ».

Par exemple, les `addHandlerAsync` méthodes (des objets [Binding](/javascript/api/office/office.binding), [CustomXMLPart](/javascript/api/office/office.customxmlpart), [document](/javascript/api/office/office.document), [RoamingSettings](/javascript/api/outlook/office.roamingsettings)et [Settings](/javascript/api/office/office.settings) ) sont utilisées pour ajouter des fonctions de gestionnaire d’événements aux éléments représentés par ces objets. Vous pouvez accéder à la `AsyncResult.value` propriété à partir de la fonction de rappel que vous transmettez à l’une des `addHandlerAsync` méthodes, mais étant donné que vous n’avez pas accès à des données ou à un objet lorsque vous ajoutez un gestionnaire d’événements, la `value` propriété renvoie toujours **undefined** si vous tentez d’y accéder.

En revanche, si vous appelez la `Document.getSelectedDataAsync` méthode, elle renvoie les données sélectionnées par l’utilisateur dans le document à la `AsyncResult.value` propriété dans le rappel. Ou, si vous appelez la méthode [bindings. getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-) , elle renvoie un tableau de tous les `Binding` objets dans le document. Si vous appelez la méthode [bindings. getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) , elle renvoie un seul `Binding` objet.

Pour obtenir une description de ce qui est renvoyé à la `AsyncResult.value` propriété pour une `Async` méthode, consultez la section « valeur de rappel » de la rubrique de référence de cette méthode. Pour obtenir un résumé de tous les objets qui fournissent des `Async` méthodes, reportez-vous au tableau en bas de la rubrique [asyncResult](/javascript/api/office/office.asyncresult) Object.


## <a name="asynchronous-programming-patterns"></a>Modèles de programmation asynchrone


L’API JavaScript pour Office prend en charge deux types de modèles de programmation asynchrone :


- Utilisation des rappels imbriqués
    
- Utilisation du modèle des promesses
    
La programmation asynchrone à l’aide des fonctions de rappel nécessite que vous imbriquiez fréquemment le résultat retourné d’un rappel au sein d’au moins deux rappels. Pour ce faire, vous pouvez utiliser les rappels imbriqués de toutes les méthodes « Async » de l’API.

L’utilisation des rappels imbriqués est un modèle de programmation familier pour la plupart des développeurs JavaScript, mais le code contenant des rappels fortement imbriqués peut être difficile à lire et à comprendre. En guise d’alternative aux rappels imbriqués, l’API JavaScript Office prend également en charge une implémentation du modèle de promesses. 

> [!NOTE]
> Dans la version actuelle de l’API JavaScript pour Office, la prise en charge *intégrée* du modèle promesses fonctionne uniquement avec le code pour les [liaisons dans les feuilles de calcul Excel et les documents Word](bind-to-regions-in-a-document-or-spreadsheet.md). Toutefois, vous pouvez encapsuler d’autres fonctions qui ont des rappels à l’intérieur de votre fonction personnalisée de renvoi de la promesse. Pour plus d’informations, consultez [la rubrique envelopper les API communes dans les fonctions de retour à la vente](#wrap-common-apis-in-promise-returning-functions).


### <a name="asynchronous-programming-using-nested-callback-functions"></a>Programmation asynchrone utilisant des fonctions de rappel imbriquées

Vous devez fréquemment effectuer au moins deux opérations asynchrones pour réaliser une tâche. Pour ce faire, vous pouvez imbriquer un appel « Async » dans un autre.

L’exemple de code suivant imbrique deux appels asynchrones.


- D’abord, la méthode [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) est appelée pour accéder à une liaison dans le document nommé « MyBinding ». L' `AsyncResult` objet renvoyé au `result` paramètre de ce rappel permet d’accéder à l’objet Binding spécifié à partir de la `AsyncResult.value` propriété.

- Ensuite, l’objet Binding auquel vous avez accédé à partir du premier `result` paramètre est utilisé pour appeler la méthode [Binding. getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-) .

- Enfin, le `result2` paramètre du rappel transmis à la `Binding.getDataAsync` méthode est utilisé pour afficher les données dans la liaison.


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

Ce modèle de rappel imbriqué de base peut être utilisé pour toutes les méthodes asynchrones dans l’API JavaScript pour Office.

Les sections suivantes montrent comment utiliser des fonctions anonymes ou nommées pour des rappels imbriqués dans des méthodes asynchrones.


#### <a name="using-anonymous-functions-for-nested-callbacks"></a>Utilisation des fonctions anonymes pour des rappels imbriqués

Dans l’exemple suivant, deux fonctions anonymes sont déclarées inline et transmises `getByIdAsync` aux `getDataAsync` méthodes et en tant que rappels imbriqués. Comme les fonctions sont très simples, l’objet de l’implémentation est évident.


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


#### <a name="using-named-functions-for-nested-callbacks"></a>Utilisation de fonctions nommées pour des rappels imbriqués

Dans des implémentations complexes, il peut être utile d’utiliser des fonctions nommées pour garantir une meilleure lisibilité, simplicité de gestion et possibilité de réutilisation du code. Dans l’exemple suivant, les deux fonctions anonymes de l’exemple de la section précédente ont été réécrites sous la forme de fonctions nommées `deleteAllData` et `showResult` . Ces fonctions nommées sont ensuite transmises aux `getByIdAsync` `deleteAllDataValuesAsync` méthodes et sous forme de rappels par nom.


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


L’API JavaScript pour Office fournit la méthode [Office. Select](/javascript/api/office#office-select-expression--callback-) pour prendre en charge le modèle de promesses pour l’utilisation d’objets Binding existants. L’objet promesse renvoyé à la `Office.select` méthode prend en charge uniquement les quatre méthodes auxquelles vous pouvez accéder directement à partir de l’objet [Binding](/javascript/api/office/office.binding) : [getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-), [setDataAsync](/javascript/api/office/office.binding#setdataasync-data--options--callback-), [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-)et [removeHandlerAsync](/javascript/api/office/office.binding#removehandlerasync-eventtype--options--callback-).


Le modèle des promesses à utiliser avec les liaisons se présente comme suit :

 **Office. Select (**_selectorExpression_, _OnError_**).** _BindingObjectAsyncMethod_

Le paramètre _selectorExpression_ prend la forme `"bindings#bindingId"` , où _bindingId_ est le nom ( `id` ) d’une liaison que vous avez créée précédemment dans le document ou la feuille de calcul (à l’aide de l’une des méthodes « addFrom » de la `Bindings` collection : `addFromNamedItemAsync` , `addFromPromptAsync` , ou `addFromSelectionAsync` ). Par exemple, l’expression de sélecteur `bindings#cities` spécifie que vous souhaitez accéder à la liaison avec l' **ID** « villes ».

Le paramètre _OnError_ est une fonction de gestion des erreurs qui accepte un seul paramètre de type `AsyncResult` qui peut être utilisé pour accéder à un `Error` objet, si la `select` méthode ne parvient pas à accéder à la liaison spécifiée. L’exemple suivant montre une fonction de gestion des erreurs de base pouvant être passée au paramètre _onError_.




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

Remplacez l’espace réservé _BindingObjectAsyncMethod_ par un appel à l’une des quatre `Binding` méthodes d’objet prises en charge par l’objet Promise : `getDataAsync` ,, `setDataAsync` `addHandlerAsync` ou `removeHandlerAsync` . Les appels à ces méthodes ne prennent pas en charge les promesses supplémentaires. Vous devez les appeler à l’aide du [modèle de fonction de rappel imbriquée](#asynchronous-programming-using-nested-callback-functions).

Une fois qu’une `Binding` promesse d’objet est satisfaite, elle peut être réutilisée dans l’appel de la méthode chaînée comme s’il s’agissait d’une liaison (le runtime du complément ne réessaie pas de manière asynchrone de répondre à la promesse). Si la `Binding` promesse de l’objet ne peut pas être satisfaite, le runtime du complément réessaiera d’accéder à l’objet Binding lors de la prochaine appel de l’une de ses méthodes asynchrones.

L’exemple de code suivant utilise la `select` méthode pour récupérer une liaison avec le `id` « `cities` » de la `Bindings` collection, puis appelle la méthode [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-) pour ajouter un gestionnaire d’événements pour l’événement [DataChanged](/javascript/api/office/office.bindingdatachangedeventargs) de la liaison.




```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```


> [!IMPORTANT]
> La `Binding` promesse de l’objet renvoyée par la `Office.select` méthode donne accès uniquement aux quatre méthodes de l' `Binding` objet. Si vous avez besoin d’accéder à n’importe quel autre membre de l' `Binding` objet, vous devez utiliser la `Document.bindings` propriété et `Bindings.getByIdAsync` ou les `Bindings.getAllAsync` méthodes pour récupérer l' `Binding` objet. Par exemple, si vous avez besoin d’accéder aux propriétés de l' `Binding` objet (les `document` `id` Propriétés, ou `type` ), ou si vous avez besoin d’accéder aux propriétés des objets [MatrixBinding](/javascript/api/office/office.matrixbinding) ou [TableBinding](/javascript/api/office/office.tablebinding) , vous devez utiliser `getByIdAsync` les `getAllAsync` méthodes ou pour récupérer un `Binding` objet.


## <a name="passing-optional-parameters-to-asynchronous-methods"></a>Passage de paramètres facultatifs à des méthodes asynchrones


La syntaxe courante pour toutes les méthodes « Async » suit ce modèle :

 _AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_`);`

Toutes les méthodes asynchrones prennent en charge des paramètres facultatifs, qui sont passés sous la forme d’un objet JSON (JavaScript Object Notation) qui contient un ou plusieurs paramètres facultatifs. L’objet JSON contenant les paramètres facultatifs est une collection non ordonnée de paires clé-valeur où le caractère « : » sépare la clé de la valeur. Chaque paire dans l’objet est séparée par une virgule, et l’ensemble complet de paires est placé entre accolades. La clé est le nom du paramètre, et la valeur est la valeur à passer pour ce paramètre.

Vous pouvez créer l’objet JSON qui contient les paramètres facultatifs incorporés, ou créer un `options` objet et le passer comme paramètre _options_ .


### <a name="passing-optional-parameters-inline"></a>Passage de paramètres facultatifs incorporés

Par exemple, la syntaxe pour appeler la méthode [Document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) avec des paramètres facultatifs incorporés se présente comme ceci :

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

Dans cette forme de syntaxe d’appel, les deux paramètres facultatifs, _coercionType_ et _asyncContext_, sont définis comme un objet incorporé entre accolades.

L’exemple suivant montre comment appeler la `Document.setSelectedDataAsync` méthode en spécifiant des paramètres facultatifs incorporés.


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


### <a name="passing-optional-parameters-in-an-options-object"></a>Passage de paramètres facultatifs dans un objet options

Vous pouvez également créer un objet nommé `options` qui spécifie les paramètres facultatifs séparément de l’appel de la méthode, puis transmettre l' `options` objet en tant qu’argument _options_ .

L’exemple suivant montre une façon de créer l' `options` objet, où `parameter1` ,, `value1` et ainsi de suite, sont des espaces réservés aux noms et valeurs de paramètres effectifs.




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

Voici une autre façon de créer l' `options` objet.




```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

Ce qui ressemble à l’exemple suivant lorsqu’il est utilisé pour spécifier les `ValueFormat` `FilterType` paramètres et :


```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```


> [!NOTE]
> Lors de l’utilisation de l’une ou l’autre des méthodes de création de l' `options` objet, vous pouvez spécifier des paramètres facultatifs dans n’importe quel ordre dans la mesure où leurs noms sont correctement spécifiés.

L’exemple suivant montre comment appeler la `Document.setSelectedDataAsync` méthode en spécifiant des paramètres facultatifs dans un `options` objet.




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


Dans les deux exemples de paramètres facultatifs, le paramètre _callback_ est spécifié en tant que dernier paramètre (en suivant les paramètres facultatifs Inline, ou en suivant l’objet d’arguments _options_ ). Vous pouvez également spécifier le paramètre _callback_ à l’intérieur de l’objet JSON incorporé, ou dans l’objet `options`. Cependant, vous ne pouvez passer le paramètre _callback_ qu’à un seul endroit : soit dans l’objet _options_ (incorporé ou créé en externe), soit comme dernier paramètre, mais pas les deux.

## <a name="wrap-common-apis-in-promise-returning-functions"></a>Envelopper les API communes dans les fonctions de retour à la vente

Les méthodes de l’API commune (et de l’API Outlook) ne retournent pas de [promesses](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise). Par conséquent, vous ne pouvez pas utiliser [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) pour suspendre l’exécution jusqu’à la fin de l’opération asynchrone. Si vous avez besoin `await` de comportement, vous pouvez encapsuler l’appel de méthode dans une promesse créée de manière explicite. 

Le modèle de base consiste à créer une méthode asynchrone qui renvoie immédiatement un objet promesse et *résout* cet objet promise au terme de la méthode interne, ou *rejette* l’objet en cas d’échec de la méthode. Voici un exemple simple

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

Lorsque cette méthode doit être attendue, elle peut être appelée soit à l’aide du `await` mot clé, soit en tant que fonction passée à une `then` fonction.

> [!NOTE]
> Cette technique est particulièrement utile lorsque vous devez appeler l’une des API communes à l’intérieur d’un appel de la `run` méthode dans l’un des modèles objet spécifiques à l’application. Pour obtenir un exemple de la fonction ci-dessus utilisée de cette façon, reportez-vous au fichier [Home.js dans l’exemple Word-Add-in-JavaScript-MDConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion/blob/master/Word-Add-in-JavaScript-MDConversionWeb/Home.js).

Voici un exemple d’utilisation de la machine à écrire.

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
