---
title: Programmation asynchrone dans des compléments Office
description: Découvrez comment la bibliothèque JavaScript Office utilise la programmation asynchrone dans Office compléments.
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7c57dc1c35d518f86e4757fb1c5d6d51c9819441
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66090949"
---
# <a name="asynchronous-programming-in-office-add-ins"></a>Programmation asynchrone dans des compléments Office

[!include[information about the common API](../includes/alert-common-api-info.md)]

Pourquoi l’API de Compléments Office a-t-elle recours à la programmation asynchrone ? JavaScript étant un langage monothread, si le script appelle un processus synchrone de longue durée, toute exécution de script ultérieure sera bloquée tant que ce processus ne sera pas terminé. Étant donné que certaines opérations sur Office clients web (mais aussi les clients riches) peuvent bloquer l’exécution si elles sont exécutées de façon synchrone, la plupart des API JavaScript Office sont conçues pour s’exécuter de façon asynchrone. Cela garantit que les compléments Office sont réactifs et rapides. Vous devez donc fréquemment écrire des fonctions de rappel lorsque vous utilisez ces méthodes asynchrones.

Les noms de toutes les méthodes asynchrones dans l’API se terminent par « Async », par exemple , `Document.getSelectedDataAsync``Binding.getDataAsync`ou `Item.loadCustomPropertiesAsync` par des méthodes. Lorsqu’une méthode « Async » est appelée, elle est exécutée immédiatement et toute exécution de script ultérieure peut se poursuivre normalement. La fonction de rappel facultative que vous transmettez à une méthode « Async » s’exécute dès que l’opération demandée ou les données sont prêtes. L’opération est généralement rapide, mais le retour pourrait présenter un léger retard.

Le diagramme suivant montre le flux d’exécution d’un appel à une méthode « Async » qui lit les données que l’utilisateur a sélectionnées dans un document ouvert dans word ou Excel basé sur le serveur. Au moment où l’appel « Async » est effectué, le thread d’exécution JavaScript est libre d’effectuer un traitement supplémentaire côté client (bien qu’aucun ne soit affiché dans le diagramme). Lorsque la méthode « Async » est retournée, le rappel reprend l’exécution sur le thread, et le complément peut accéder aux données, y faire quelque chose et afficher le résultat. Le même modèle d’exécution asynchrone s’applique lors de l’utilisation des Office applications clientes riches, telles que Word 2013 ou Excel 2013.

*Figure 1. Flux d’exécution de programmation asynchrone*

![Diagramme montrant l’interaction de l’exécution de commande au fil du temps avec l’utilisateur, la page de complément et le serveur d’applications web hébergeant le complément.](../images/office-addins-asynchronous-programming-flow.png)

La prise en charge de cette conception asynchrone dans les clients riches et les clients web fait partie des objectifs de conception « écriture unique-exécution multiplateforme » du modèle de développement des Compléments Office. Par exemple, vous pouvez créer un complément de contenu ou du volet de tâches avec une seule base de code qui sera exécutée sur Excel 2013 et Excel sur le web.

## <a name="write-the-callback-function-for-an-async-method"></a>Écrire la fonction de rappel pour une méthode « Async »

La fonction de rappel que vous transmettez en tant qu’argument de _rappel_ à une méthode « Async » doit déclarer un paramètre unique que le runtime de complément utilisera pour fournir l’accès à un objet [AsyncResult](/javascript/api/office/office.asyncresult) lorsque la fonction de rappel s’exécute. Vous pouvez écrire:

- Fonction anonyme qui doit être écrite et passée directement en ligne avec l’appel à la méthode « Async » comme paramètre de _rappel_ de la méthode « Async ».

- Fonction nommée, en passant le nom de cette fonction comme paramètre de _rappel_ d’une méthode « Async ».

Une fonction anonyme est utile si vous envisagez de n’utiliser son code qu’une fois : comme elle n’a pas de nom, vous ne pouvez pas y faire référence dans une autre partie du code. Une fonction nommée est utile si vous voulez réutiliser la fonction de rappel pour plusieurs méthodes « Async ».

### <a name="write-an-anonymous-callback-function"></a>Écrire une fonction de rappel anonyme

La fonction de rappel anonyme suivante déclare un seul paramètre nommé `result` qui récupère des données à partir de la propriété [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member) lorsque le rappel est retourné.

```js
function (result) {
        write('Selected data: ' + result.value);
}
```

L’exemple suivant montre comment passer cette fonction de rappel anonyme en ligne dans le contexte d’un appel de méthode « Async » complet à la `Document.getSelectedDataAsync` méthode.

- Le premier argument _coercionType_ , `Office.CoercionType.Text`spécifie de retourner les données sélectionnées sous la forme d’une chaîne de texte.

- Le deuxième argument _de rappel_ est la fonction anonyme passée en ligne à la méthode. Lorsque la fonction s’exécute, elle utilise le paramètre de _résultat_ pour accéder à la `value` propriété de l’objet `AsyncResult` afin d’afficher les données sélectionnées par l’utilisateur dans le document.

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

Vous pouvez également utiliser le paramètre de votre fonction de rappel pour accéder à d’autres propriétés de l’objet `AsyncResult` . Utilisez la propriété [AsyncResult.status](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) pour déterminer si l’appel a réussi ou échoué. En cas d’échec, vous pouvez utiliser la propriété [AsyncResult.error](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) pour accéder à un objet [Error](/javascript/api/office/office.error) et obtenir des informations sur l’erreur.

Pour plus d’informations sur l’utilisation de la `getSelectedDataAsync` méthode, consultez [Lire et écrire des données dans la sélection active d’un document ou d’une feuille de calcul](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md). 

### <a name="write-a-named-callback-function"></a>Écrire une fonction de rappel nommée

Vous pouvez également écrire une fonction nommée et passer son nom au paramètre de _rappel_ d’une méthode « Async ». Par exemple, l’exemple précédent peut être réécrit pour passer une fonction nommée `writeDataCallback` en tant que paramètre _callback_ comme suit.

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

Le `asyncContext`, `status`et `error` les propriétés de l’objet `AsyncResult` retournent les mêmes types d’informations à la fonction de rappel passée à toutes les méthodes « Async ». Toutefois, ce qui est retourné à la `AsyncResult.value` propriété varie en fonction des fonctionnalités de la méthode « Async ».

Par exemple, les `addHandlerAsync` méthodes (des objets [Binding](/javascript/api/office/office.binding), [CustomXmlPart](/javascript/api/office/office.customxmlpart), [Document](/javascript/api/office/office.document), [RoamingSettings](/javascript/api/outlook/office.roamingsettings) et [Paramètres](/javascript/api/office/office.settings)) sont utilisées pour ajouter des fonctions de gestionnaire d’événements aux éléments représentés par ces objets. Vous pouvez accéder à la `AsyncResult.value` propriété à partir de la fonction de rappel que vous transmettez à l’une `addHandlerAsync` des méthodes, mais étant donné qu’aucune donnée ou objet n’est accessible lorsque vous ajoutez un gestionnaire d’événements, la `value` propriété retourne toujours **une valeur non définie** si vous tentez d’y accéder.

En revanche, si vous appelez la `Document.getSelectedDataAsync` méthode, elle renvoie les données sélectionnées par l’utilisateur dans le document à la `AsyncResult.value` propriété dans le rappel. Ou, si vous appelez la méthode [Bindings.getAllAsync](/javascript/api/office/office.bindings#office-office-bindings-getallasync-member(1)) , elle retourne un tableau de `Binding` tous les objets du document. Et, si vous appelez la méthode [Bindings.getByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-getbyidasync-member(1)) , elle renvoie un seul `Binding` objet.

Pour obtenir une description de ce qui est retourné à la `AsyncResult.value` propriété pour une `Async` méthode, consultez la section « Valeur de rappel » de la rubrique de référence de cette méthode. Pour obtenir un résumé de tous les objets qui fournissent `Async` des méthodes, consultez le tableau en bas de la rubrique de l’objet [AsyncResult](/javascript/api/office/office.asyncresult) .

## <a name="asynchronous-programming-patterns"></a>Modèles de programmation asynchrone

L’API JavaScript Office prend en charge deux types de modèles de programmation asynchrones.

- Utilisation des rappels imbriqués
- Utilisation du modèle des promesses

La programmation asynchrone à l’aide des fonctions de rappel nécessite que vous imbriquiez fréquemment le résultat retourné d’un rappel au sein d’au moins deux rappels. Pour ce faire, vous pouvez utiliser les rappels imbriqués de toutes les méthodes « Async » de l’API.

L’utilisation des rappels imbriqués est un modèle de programmation familier pour la plupart des développeurs JavaScript, mais le code contenant des rappels fortement imbriqués peut être difficile à lire et à comprendre. En guise d’alternative aux rappels imbriqués, l’API JavaScript Office prend également en charge une implémentation du modèle de promesses.

> [!NOTE]
> Dans la version actuelle de l’API JavaScript Office, la prise *en charge intégrée* du modèle de promesses fonctionne uniquement avec du code pour [les liaisons dans les feuilles de calcul Excel et les documents Word](bind-to-regions-in-a-document-or-spreadsheet.md). Toutefois, vous pouvez encapsuler d’autres fonctions qui ont des rappels à l’intérieur de votre propre fonction de retour de promesse personnalisée. Pour plus d’informations, consultez [Wrap Common API in Promise-returning functions](#wrap-common-apis-in-promise-returning-functions).

### <a name="asynchronous-programming-using-nested-callback-functions"></a>Programmation asynchrone utilisant des fonctions de rappel imbriquées

Vous devez fréquemment effectuer au moins deux opérations asynchrones pour réaliser une tâche. Pour ce faire, vous pouvez imbriquer un appel « Async » dans un autre.

L’exemple de code suivant imbrique deux appels asynchrones.

- D’abord, la méthode [Bindings.getByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-getbyidasync-member(1)) est appelée pour accéder à une liaison dans le document nommé « MyBinding ». L’objet `AsyncResult` retourné au `result` paramètre de ce rappel fournit l’accès à l’objet de liaison spécifié à partir de la `AsyncResult.value` propriété.
- Ensuite, l’objet de liaison accessible à partir du premier `result` paramètre est utilisé pour appeler la méthode [Binding.getDataAsync](/javascript/api/office/office.binding#office-office-binding-getdataasync-member(1)) .
- Enfin, le `result2` paramètre du rappel passé à la `Binding.getDataAsync` méthode est utilisé pour afficher les données dans la liaison.

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

Ce modèle de rappel imbriqué de base peut être utilisé pour toutes les méthodes asynchrones dans l’API JavaScript Office.

Les sections suivantes montrent comment utiliser des fonctions anonymes ou nommées pour des rappels imbriqués dans des méthodes asynchrones.

#### <a name="use-anonymous-functions-for-nested-callbacks"></a>Utiliser des fonctions anonymes pour les rappels imbriqués

Dans l’exemple suivant, deux fonctions anonymes sont déclarées inline et transmises dans les `getByIdAsync` méthodes et `getDataAsync` les rappels imbriqués. Comme les fonctions sont très simples, l’objet de l’implémentation est évident.

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

#### <a name="use-named-functions-for-nested-callbacks"></a>Utiliser des fonctions nommées pour les rappels imbriqués

Dans des implémentations complexes, il peut être utile d’utiliser des fonctions nommées pour garantir une meilleure lisibilité, simplicité de gestion et possibilité de réutilisation du code. Dans l’exemple suivant, les deux fonctions anonymes de l’exemple de la section précédente ont été réécrites en tant que fonctions nommées `deleteAllData` et `showResult`. Ces fonctions nommées sont ensuite passées dans les `getByIdAsync` méthodes et `deleteAllDataValuesAsync` les rappels par nom.

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

L’API JavaScript Office fournit la méthode [Office.select](/javascript/api/office#Office_select_expression__callback_) pour prendre en charge le modèle de promesses pour l’utilisation d’objets de liaison existants. L’objet promise retourné à la `Office.select` méthode prend en charge uniquement les quatre méthodes auxquelles vous pouvez accéder directement à partir de l’objet [Binding](/javascript/api/office/office.binding) : [getDataAsync](/javascript/api/office/office.binding#office-office-binding-getdataasync-member(1)), [setDataAsync](/javascript/api/office/office.binding#office-office-binding-setdataasync-member(1)), [addHandlerAsync](/javascript/api/office/office.binding#office-office-binding-addhandlerasync-member(1)) et [removeHandlerAsync](/javascript/api/office/office.binding#office-office-binding-removehandlerasync-member(1)).

Le modèle de promesses pour l’utilisation des liaisons prend cette forme.

**Office.select(**_selectorExpression_, _onError_**).** _BindingObjectAsyncMethod_

Le paramètre _selectorExpression_ prend la forme `"bindings#bindingId"`, où _bindingId_ est le nom ( `id`) d’une liaison que vous avez créée précédemment dans le document ou la feuille de calcul (à l’aide de l’une des méthodes « addFrom » de la `Bindings` collection : `addFromNamedItemAsync`, `addFromPromptAsync`ou `addFromSelectionAsync`). Par exemple, l’expression `bindings#cities` de sélecteur spécifie que vous souhaitez accéder à la liaison avec un **ID** « cities ».

Le paramètre _onError_ est une fonction de gestion des erreurs qui prend un seul paramètre de type `AsyncResult` qui peut être utilisé pour accéder à un `Error` objet, si la `select` méthode ne parvient pas à accéder à la liaison spécifiée. L’exemple suivant montre une fonction de gestion des erreurs de base pouvant être passée au paramètre _onError_.

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

Remplacez l’espace réservé _BindingObjectAsyncMethod_ par un appel à l’une des quatre `Binding` méthodes d’objet prises en charge par l’objet promise : `getDataAsync`, `setDataAsync`, `addHandlerAsync`ou `removeHandlerAsync`. Les appels à ces méthodes ne prennent pas en charge les promesses supplémentaires. Vous devez les appeler à l’aide du [modèle de fonction de rappel imbriquée](#asynchronous-programming-using-nested-callback-functions).

Une fois qu’une `Binding` promesse d’objet est remplie, elle peut être réutilisée dans l’appel de méthode chaîné comme s’il s’agissait d’une liaison (le runtime de complément ne réessayera pas de manière asynchrone de remplir la promesse). Si la `Binding` promesse d’objet ne peut pas être satisfaite, le runtime de complément tente à nouveau d’accéder à l’objet de liaison la prochaine fois qu’une de ses méthodes asynchrones est appelée.

L’exemple de code suivant utilise la `select` méthode pour récupérer une liaison avec le `id` «`cities` » de la `Bindings` collection, puis appelle la méthode [addHandlerAsync](/javascript/api/office/office.binding#office-office-binding-addhandlerasync-member(1)) pour ajouter un gestionnaire d’événements pour l’événement [dataChanged](/javascript/api/office/office.bindingdatachangedeventargs) de la liaison.

```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```

> [!IMPORTANT]
> La `Binding` promesse d’objet retournée par la `Office.select` méthode permet d’accéder uniquement aux quatre méthodes de l’objet `Binding` . Si vous devez accéder à l’un des autres membres de l’objet `Binding` , vous devez utiliser la `Document.bindings` propriété et `Bindings.getByIdAsync` ou `Bindings.getAllAsync` les méthodes pour récupérer l’objet `Binding` . Par exemple, si vous devez accéder à l’une des propriétés de l’objet `Binding` (les `document`, `id`ou `type` propriétés) ou accéder aux propriétés des objets [MatrixBinding](/javascript/api/office/office.matrixbinding) ou [TableBinding](/javascript/api/office/office.tablebinding) , vous devez utiliser la `getByIdAsync` ou `getAllAsync` les méthodes pour récupérer un `Binding` objet.

## <a name="pass-optional-parameters-to-asynchronous-methods"></a>Passer des paramètres facultatifs aux méthodes asynchrones

La syntaxe commune pour toutes les méthodes « Async » suit ce modèle.

 _AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_`);`

Toutes les méthodes asynchrones prennent en charge les paramètres facultatifs, qui sont passés en tant qu’objet JavaScript qui contient un ou plusieurs paramètres facultatifs. L’objet contenant les paramètres facultatifs est une collection non triée de paires clé-valeur avec le caractère « : » séparant la clé et la valeur. Chaque paire dans l’objet est séparée par une virgule, et l’ensemble complet de paires est placé entre accolades. La clé est le nom du paramètre, et la valeur est la valeur à passer pour ce paramètre.

Vous pouvez créer l’objet qui contient des paramètres facultatifs inline, ou en créant un `options` objet et en le transmettant comme paramètre _d’options_ .

### <a name="pass-optional-parameters-inline"></a>Passer des paramètres facultatifs inline

Par exemple, la syntaxe pour appeler la méthode [Document.setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) avec des paramètres facultatifs incorporés se présente comme ceci :

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

Sous cette forme de syntaxe appelante, les deux paramètres facultatifs, _coercionType_ et _asyncContext_, sont définis comme un objet JavaScript anonyme inclus entre accolades.

L’exemple suivant montre comment appeler la `Document.setSelectedDataAsync` méthode en spécifiant des paramètres facultatifs inline.

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
> Vous pouvez spécifier des paramètres facultatifs dans n’importe quel ordre dans l’objet de paramètre tant que leurs noms sont spécifiés correctement.

### <a name="pass-optional-parameters-in-an-options-object"></a>Passer des paramètres facultatifs dans un objet options

Vous pouvez également créer un objet nommé `options` qui spécifie les paramètres facultatifs séparément de l’appel de méthode, puis passer l’objet `options` comme argument _d’options_ .

L’exemple suivant montre une façon de créer l’objet `options` , où `parameter1`, `value1`et ainsi de suite, sont des espaces réservés pour les noms et valeurs de paramètres réels.

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

Voici une autre façon de créer l’objet `options` .

```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

Qui ressemble à l’exemple suivant lorsqu’il est utilisé pour spécifier les paramètres et `FilterType` les `ValueFormat` éléments suivants :

```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```

> [!NOTE]
> Lorsque vous utilisez l’une des méthodes de création de l’objet `options` , vous pouvez spécifier des paramètres facultatifs dans n’importe quel ordre, tant que leurs noms sont spécifiés correctement.

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

Dans les deux exemples de paramètres facultatifs, le paramètre _de rappel_ est spécifié comme dernier paramètre (en suivant les paramètres facultatifs inclus ou en suivant l’objet d’argument _options_ ). Vous pouvez également spécifier le paramètre _de rappel_ à l’intérieur de l’objet JavaScript inline ou dans l’objet `options` . Cependant, vous ne pouvez passer le paramètre _callback_ qu’à un seul endroit : soit dans l’objet _options_ (incorporé ou créé en externe), soit comme dernier paramètre, mais pas les deux.

## <a name="wrap-common-apis-in-promise-returning-functions"></a>Inclure des API courantes dans des fonctions de retour de promesse

Les méthodes d’API communes (et d’API Outlook) ne retournent pas [promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise). Par conséquent, vous ne pouvez pas utiliser [Await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) pour suspendre l’exécution tant que l’opération asynchrone n’est pas terminée. Si vous avez besoin d’un `await` comportement, vous pouvez inclure l’appel de méthode dans une promesse créée explicitement. 

Le modèle de base consiste à créer une méthode asynchrone qui retourne immédiatement un objet Promise et *résout* cet objet Promise lorsque la méthode interne se termine ou *rejette* l’objet en cas d’échec de la méthode. Voici un exemple simple.

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

Lorsque cette méthode doit être attendue, elle peut être appelée avec le `await` mot clé ou en tant que fonction passée à une `then` fonction.

> [!NOTE]
> Cette technique est particulièrement utile lorsque vous devez appeler l’une des API communes à l’intérieur d’un appel de la `run` méthode dans l’un des modèles objet spécifiques à l’application. Pour obtenir un exemple de la fonction ci-dessus utilisée de cette façon, consultez le fichier [Home.js dans l’exemple Word-Add-in-JavaScript-MDConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion/blob/master/Word-Add-in-JavaScript-MDConversionWeb/Home.js).

Voici un exemple d’utilisation de TypeScript.

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
