---
title: Programmation asynchrone dans des compléments Office
description: ''
ms.date: 01/14/2020
localization_priority: Priority
ms.openlocfilehash: 009f8e37cc8a6eb2e808278df88f3bfdc5b0d1b1
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217241"
---
# <a name="asynchronous-programming-in-office-add-ins"></a>Programmation asynchrone dans des compléments Office

Pourquoi l’API de Compléments Office a-t-elle recours à la programmation asynchrone ?JavaScript étant un langage monothread, si le script appelle un processus synchrone de longue durée, toute exécution de script ultérieure sera bloquée tant que ce processus ne sera pas terminé. Comme certaines opérations, notamment celles agissant sur les clients web Office (mais aussi sur les clients riches), peuvent bloquer l’exécution si elles sont exécutées de façon synchrone, la plupart des méthodes dans l’interface API JavaScript pour Office sont conçues pour être exécutées de façon asynchrone. Cela permet de garantir que les Compléments Office sont réactifs et très performants. Vous devez donc fréquemment écrire des fonctions de rappel lorsque vous utilisez ces méthodes asynchrones.

Le nom de toutes les méthodes asynchrones de l’API se terminent par « Async », comme pour les méthodes `Document.getSelectedDataAsync`, `Binding.getDataAsync` ou `Item.loadCustomPropertiesAsync`. Lorsqu’une méthode « Async » est appelée, elle est exécutée immédiatement et toute exécution de script ultérieure peut se poursuivre normalement. La fonction de rappel facultative que vous transmettez à une méthode « Async » s’exécute dès que l’opération demandée ou les données sont prêtes. L’opération est généralement rapide, mais le retour pourrait présenter un léger retard.

Le diagramme suivant présente le flux d’exécution d’un appel à une méthode « Async » qui lit les données sélectionnées par l’utilisateur dans un document ouvert dans l’instance Word ou Excel sur le serveur. Au moment où l’appel « Async » est effectué, le thread d’exécution JavaScript est libre d’effectuer tout traitement côté client supplémentaire (même si aucun n’est affiché dans le diagramme). Lors du retour de la méthode « Async », l’appel reprend l’exécution sur le thread et le complément peut accéder aux données, les exploiter et afficher le résultat. Le même motif d’exécution asynchrone est employé en cas d’utilisation des applications hôtes de client riche Office, telles que Word 2013 ou Excel 2013.

*Figure 1. Flux d’exécution de programmation asynchrone*

![Flux d’exécution de thread de programmation asynchrone](../images/office-addins-asynchronous-programming-flow.png)

La prise en charge de cette conception asynchrone dans les clients riches et les clients web fait partie des objectifs de conception « écriture unique-exécution multiplateforme » du modèle de développement des Compléments Office. Par exemple, vous pouvez créer un complément de contenu ou du volet de tâches avec une seule base de code qui sera exécutée sur Excel 2013 et Excel sur le web.

## <a name="writing-the-callback-function-for-an-async-method"></a>Écriture de la fonction de rappel pour une méthode « Async »


La fonction de rappel que vous transmettez en tant qu’argument _callback_ à une méthode « Async » doit déclarer un seul paramètre que le runtime de complément va utiliser pour permettre l’accès à un objet [AsyncResult](/javascript/api/office/office.asyncresult) lorsque la fonction de rappel sera exécutée. Vous pouvez écrire :


- une fonction anonyme devant être écrite et passée directement en ligne avec l’appel à la méthode « Async » en tant que paramètre  _callback_ de la méthode « Async » ;

- une fonction nommée, en passant le nom de cette fonction en tant que paramètre  _callback_ de la méthode « Async ».

Une fonction anonyme est utile si vous envisagez de n’utiliser son code qu’une fois : comme elle n’a pas de nom, vous ne pouvez pas y faire référence dans une autre partie du code. Une fonction nommée est utile si vous voulez réutiliser la fonction de rappel pour plusieurs méthodes « Async ».


### <a name="writing-an-anonymous-callback-function"></a>Écriture d’une fonction de rappel anonyme

La fonction de rappel anonyme suivante déclare un seul paramètre nommé `result` qui récupère les données à partir de la propriété [AsyncResult.value](/javascript/api/office/office.asyncresult#value) lorsque le rappel est renvoyé.


```js
function (result) {
        write('Selected data: ' + result.value);
}
```

L’exemple suivant montre comment passer cette fonction de rappel anonyme dans le contexte d’un appel complet de méthode « Async » à la méthode  **Document.getSelectedDataAsync**.


- Le premier argument  _coercionType_,  `Office.CoercionType.Text`, spécifie le retour des données sélectionnées en tant que chaîne de texte.

- Le deuxième argument  _callback_ est la fonction anonyme passée en ligne à la méthode. Lors de l’exécution de la fonction, elle utilise le paramètre _result_ pour accéder à la propriété **value** de l’objet **AsyncResult** afin d’afficher les données sélectionnées par l’utilisateur dans le document.


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

Vous pouvez également utiliser le paramètre de votre fonction de rappel pour accéder aux autres propriétés de l’objet **AsyncResult**. Utilisez la propriété [AsyncResult.status](/javascript/api/office/office.asyncresult#status) pour déterminer si l’appel a réussi ou échoué. En cas d’échec, vous pouvez utiliser la propriété [AsyncResult.error](/javascript/api/office/office.asyncresult#error) pour accéder à un objet [Error](/javascript/api/office/office.error) et obtenir des informations sur l’erreur.

Pour plus d’informations sur l’utilisation de la méthode  **getSelectedDataAsync**, voir [Lecture et écriture de données dans la sélection active d’un document ou d’une feuille de calcul](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md). 


### <a name="writing-a-named-callback-function"></a>Écriture d’une fonction de rappel nommée

Vous pouvez également écrire une fonction nommée et passer son nom au paramètre  _callback_ d’une méthode « Async ». Par exemple, l’exemple précédent peut être réécrit pour passer une fonction nommée `writeDataCallback` en tant que paramètre _callback_ comme suit.


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


Les propriétés  **asyncContext**,  **status** et **error** de l’objet **AsyncResult** retournent le même type d’informations à la fonction de rappel passée à toutes les méthodes « Async ». Cependant, les éléments retournés à la propriété **AsyncResult.value** varient selon la fonctionnalité de la méthode « Async ».

Par exemple, les méthodes **addHandlerAsync** (des objets [Binding](/javascript/api/office/office.binding), [CustomXmlPart](/javascript/api/office/office.customxmlpart), [Document](/javascript/api/office/office.document), [RoamingSettings](/javascript/api/outlook/office.roamingsettings) et [Settings](/javascript/api/office/office.settings)) sont utilisées pour ajouter des fonctions de gestionnaire d’événements aux éléments représentés par ces objets. Vous pouvez accéder à la propriété **AsyncResult.value** à partir de la fonction de rappel que vous transmettez aux méthodes **addHandlerAsync**, mais comme vous n’accédez à aucune donnée ni à aucun objet lorsque vous ajoutez un gestionnaire d’événements, la propriété **value** renvoie toujours **undefined** si vous tentez d’y accéder.

En revanche, si vous appelez la méthode  **Document.getSelectedDataAsync**, celle-ci renvoie les données que l’utilisateur a sélectionnées dans le document à la propriété  **AsyncResult.value** dans le rappel. Ou alors, si vous appelez la méthode [Bindings.getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-), celle-ci renvoie un tableau de tous les objets  **Binding** du document. Enfin, si vous appelez la méthode [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-), celle-ci renvoie un seul objet  **Binding**.

Pour obtenir une description des éléments renvoyés à la propriété **AsyncResult.value** pour une méthode « Async », voir la section relative à la valeur de rappel dans la rubrique de référence de cette méthode. Pour obtenir un résumé de tous les objets qui fournissent des méthodes « Async », voir le tableau situé au bas de la rubrique relative à l’objet [AsyncResult](/javascript/api/office/office.asyncresult).


## <a name="asynchronous-programming-patterns"></a>Modèles de programmation asynchrone


L’interface API JavaScript pour Office prend en charge deux types de modèles de programmation asynchrone :


- Utilisation des rappels imbriqués
    
- Utilisation du modèle des promesses
    
La programmation asynchrone à l’aide des fonctions de rappel nécessite que vous imbriquiez fréquemment le résultat retourné d’un rappel au sein d’au moins deux rappels. Pour ce faire, vous pouvez utiliser les rappels imbriqués de toutes les méthodes « Async » de l’API.

L’utilisation des rappels imbriqués est un modèle de programmation familier pour la plupart des développeurs JavaScript, mais le code contenant des rappels fortement imbriqués peut être difficile à lire et à comprendre. Pour offrir une solution de remplacement aux rappels imbriqués, l’interface API JavaScript pour Office prend également en charge l’implémentation du modèle des promesses. Cependant, dans la version actuelle de l’interface API JavaScript pour Office, le modèle des promesses fonctionne uniquement avec du code destiné aux [liaisons dans les feuilles de calcul Excel et les documents Word](bind-to-regions-in-a-document-or-spreadsheet.md).

<a name="AsyncProgramming_NestedCallbacks" />
### <a name="asynchronous-programming-using-nested-callback-functions"></a>Programmation asynchrone utilisant des fonctions de rappel imbriquées


Vous devez fréquemment effectuer au moins deux opérations asynchrones pour réaliser une tâche. Pour ce faire, vous pouvez imbriquer un appel « Async » dans un autre.

L’exemple de code suivant imbrique deux appels asynchrones.


- D’abord, la méthode [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) est appelée pour accéder à une liaison dans le document nommé « MyBinding ». L’objet **AsyncResult** renvoyé au paramètre `result` de ce rappel donne accès à l’objet de liaison spécifié dans la propriété **AsyncResult.value**.

- Ensuite, l’objet Binding auquel vous avez accédé à partir du premier paramètre `result` est utilisé pour appeler la méthode [Binding.getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-).

- Enfin, le paramètre  `result2` du rappel passé à la méthode **Binding.getDataAsync** est utilisé pour afficher les données dans la liaison.


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

Ce modèle de rappel imbriqué de base s’applique à toutes les méthodes asynchrones dans l’interface API JavaScript pour Office.

Les sections suivantes montrent comment utiliser des fonctions anonymes ou nommées pour des rappels imbriqués dans des méthodes asynchrones.


#### <a name="using-anonymous-functions-for-nested-callbacks"></a>Utilisation des fonctions anonymes pour des rappels imbriqués

Dans l’exemple suivant, deux fonctions anonymes sont déclarées en ligne et passées dans les méthodes  **getByIdAsync** et **getDataAsync** en tant que rappels imbriqués. Comme les fonctions sont très simples, l’objet de l’implémentation est évident.


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

Dans des implémentations complexes, il peut être utile d’utiliser des fonctions nommées pour garantir une meilleure lisibilité, simplicité de gestion et possibilité de réutilisation du code. Dans l’exemple suivant, les deux fonctions anonymes de l’exemple dans la section précédente ont été réécrites comme fonctions nommées  `deleteAllData` et `showResult`. Ces fonctions nommées sont ensuite passées dans les méthodes  **getByIdAsync** et **deleteAllDataValuesAsync** comme rappels par nom.


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

L’interface API JavaScript pour Office fournit la méthode [Office.select](/javascript/api/office#office-select-expression--callback-) pour prendre en charge le modèle des promesses permettant d’utiliser des objets de liaison existants. L’objet de promesse renvoyé à la méthode **Office.select** prend en charge uniquement les quatre méthodes auxquelles vous pouvez accéder directement à partir de l’objet [Binding](/javascript/api/office/office.binding) : [getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-), [setDataAsync](/javascript/api/office/office.binding#setdataasync-data--options--callback-), [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-) et [removeHandlerAsync](/javascript/api/office/office.binding#removehandlerasync-eventtype--options--callback-).

Le modèle des promesses à utiliser avec les liaisons se présente comme suit :

 **Office.select(**_selectorExpression_,  _onError_**).**_BindingObjectAsyncMethod_

Le paramètre  _selectorExpression_ a le format `"bindings#bindingId"`, où  _bindingId_ est le nom ( **id**) d’une liaison créée précédemment dans le document ou la feuille de calcul (à l’aide de l’une des méthodes « addFrom » de la collection  **Bindings** :  **addFromNamedItemAsync**,  **addFromPromptAsync** ou **addFromSelectionAsync**). Par exemple, l’expression de sélecteur  `bindings#cities` spécifie que vous voulez accéder à la liaison avec le paramètre **id** 'cities'.

Le paramètre  _onError_ est une fonction de gestion des erreurs qui prend un seul paramètre de type **AsyncResult** pouvant être utilisé pour accéder à un objet **Error** si la méthode **select** ne permet pas d’accéder à la liaison spécifiée. L’exemple suivant montre une fonction de gestion des erreurs de base pouvant être passée au paramètre _onError_.




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

Remplacez l’espace réservé _BindingObjectAsyncMethod_ par un appel à l’une des quatre méthodes d’objet **Binding** prises en charge par l’objet de promesse : **getDataAsync**, **setDataAsync**, **addHandlerAsync** ou **removeHandlerAsync**. Les appels à ces méthodes ne prennent pas en charge les promesses supplémentaires. Vous devez les appeler à l’aide du [modèle de fonction de rappel imbriquée](#AsyncProgramming_NestedCallbacks).

Une fois qu’une promesse d’objet  **Binding** est concrétisée, elle peut être réutilisée dans l’appel de méthode chaîné comme s’il s’agissait d’une liaison (le runtime de complément ne retentera pas de concrétiser la promesse de façon asynchrone). Si la promesse d’objet **Binding** ne peut pas être concrétisée, le runtime de complément retentera d’accéder à l’objet de liaison au prochain appel de l’une de ses méthodes asynchrones.

L’exemple de code suivant utilise la méthode **select** pour récupérer une liaison avec l’**id** « `cities` » à partir de la collection **Bindings**, puis appelle la méthode [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-) afin d’ajouter un gestionnaire d’événements pour l’événement [dataChanged](/javascript/api/office/office.bindingdatachangedeventargs) de la liaison.




```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```


> [!IMPORTANT]
> La promesse d’objet **Binding** renvoyée par la méthode **Office.select** fournit uniquement un accès aux quatre méthodes de l’objet **Binding**. Pour accéder à l’un des autres membres de l’objet **Binding**, vous devez utiliser la propriété **Document.bindings** et la méthode **Bindings.getByIdAsync** ou **Bindings.getAllAsync** pour récupérer l’objet **Binding**. Par exemple, pour accéder aux propriétés de l’objet **Binding** (propriété **document**, **id** ou **type**) ou pour accéder aux propriétés de l’objet [MatrixBinding](/javascript/api/office/office.matrixbinding) ou [TableBinding](/javascript/api/office/office.tablebinding), vous devez utiliser la méthode **getByIdAsync** ou **getAllAsync** pour récupérer un objet **Binding**.


## <a name="passing-optional-parameters-to-asynchronous-methods"></a>Passage de paramètres facultatifs à des méthodes asynchrones


La syntaxe courante pour toutes les méthodes « Async » suit ce modèle :

 _AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_`);`

Toutes les méthodes asynchrones prennent en charge des paramètres facultatifs, qui sont passés sous la forme d’un objet JSON (JavaScript Object Notation) qui contient un ou plusieurs paramètres facultatifs. L’objet JSON contenant les paramètres facultatifs est une collection non ordonnée de paires clé-valeur où le caractère « : » sépare la clé de la valeur. Chaque paire dans l’objet est séparée par une virgule, et l’ensemble complet de paires est placé entre accolades. La clé est le nom du paramètre, et la valeur est la valeur à passer pour ce paramètre.

Vous pouvez créer l’objet JSON qui contient les paramètres facultatifs incorporés, ou créer un objet  `options` et le passer comme paramètre _options_.


### <a name="passing-optional-parameters-inline"></a>Passage de paramètres facultatifs incorporés

Par exemple, la syntaxe pour appeler la méthode [Document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) avec des paramètres facultatifs incorporés se présente comme ceci :

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

Dans cette forme de syntaxe d’appel, les deux paramètres facultatifs,  _coercionType_ et _asyncContext_, sont définis comme un objet incorporé mis entre accolades.

L’exemple suivant montre comment appeler la méthode **Document.setSelectedDataAsync** en spécifiant des paramètres facultatifs incorporés.


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

Vous pouvez également créer un objet nommé  `options` qui spécifie les paramètres facultatifs séparément de l’appel de la méthode, puis passe l’objet `options` comme l’argument _options_.

L’exemple suivant illustre une manière de créer l’objet  `options`, où  `parameter1` et `value1` notamment sont des espaces réservés aux noms et valeurs de paramètres effectifs.




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

Voici une autre façon de créer l’objet  `options`.




```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

Ce qui ressemble à l’exemple suivant lors de la spécification des paramètres  **ValueFormat** et **FilterType** :


```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```


> [!NOTE]
> Au moment de créer l’objet `options` en employant l’une ou l’autre de ces méthodes, vous pouvez spécifier des paramètres facultatifs dans n’importe quel ordre du moment où leurs noms sont spécifiés correctement.

L’exemple suivant illustre comment appeler la méthode **Document.setSelectedDataAsync** en spécifiant des paramètres facultatifs dans un objet `options`.




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


Dans les deux exemples de paramètres facultatifs, le paramètre _callback_ est spécifié comme le dernier paramètre (à la suite des paramètres facultatifs incorporés, ou de l’objet de l’argument _options_). Vous pouvez également spécifier le paramètre _callback_ à l’intérieur de l’objet JSON incorporé, ou dans l’objet `options`. Cependant, vous ne pouvez passer le paramètre _callback_ qu’à un seul endroit : soit dans l’objet _options_ (incorporé ou créé en externe), soit comme dernier paramètre, mais pas les deux.


## <a name="see-also"></a>Voir aussi

- [Présentation de l’API JavaScript pour Office](understanding-the-javascript-api-for-office.md)
- [Interface API JavaScript pour Office](/office/dev/add-ins/reference/javascript-api-for-office)
