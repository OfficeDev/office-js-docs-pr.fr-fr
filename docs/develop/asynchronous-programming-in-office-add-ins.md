---
title: Programmation asynchrone dans des compl?ments Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: d251ebfd03227569b9a24bcd7f17baada6099938
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="asynchronous-programming-in-office-add-ins"></a>Programmation asynchrone dans des compl?ments Office

Pourquoi l?API de Compl?ments Office a-t-elle recours ? la programmation asynchrone ?JavaScript ?tant un langage monothread, si le script appelle un processus synchrone de longue dur?e, toute ex?cution de script ult?rieure sera bloqu?e tant que ce processus ne sera pas termin?. Comme certaines op?rations, notamment celles agissant sur les clients web Office (mais aussi sur les clients riches), peuvent bloquer l?ex?cution si elles sont ex?cut?es de fa?on synchrone, la plupart des m?thodes dans l?interface API JavaScript pour Office sont con?ues pour ?tre ex?cut?es de fa?on asynchrone. Cela permet de garantir que les Compl?ments Office sont r?actifs et tr?s performants. Vous devez donc fr?quemment ?crire des fonctions de rappel lorsque vous utilisez ces m?thodes asynchrones.

Le nom de toutes les m?thodes asynchrones de l?API se terminent par ? Async ?, comme pour les m?thodes [Document.getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync), [Binding.getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync) ou [Item.loadCustomPropertiesAsync](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item). Lorsqu?une m?thode ? Async ? est appel?e, elle est ex?cut?e imm?diatement et toute ex?cution de script ult?rieure peut se poursuivre normalement. La fonction de rappel facultative que vous transmettez ? une m?thode ? Async ? s?ex?cute d?s que l?op?ration demand?e ou les donn?es sont pr?tes. L?op?ration est g?n?ralement rapide, mais le retour pourrait pr?senter un l?ger retard.

Le diagramme suivant pr?sente le flux d?ex?cution d?un appel ? une m?thode ? Async ? qui lit les donn?es s?lectionn?es par l?utilisateur dans un document ouvert dans l?instance Word Online ou Excel Online sur le serveur. Au moment o? l?appel ? Async ? est effectu?, le thread d?ex?cution JavaScript est libre d?effectuer tout traitement c?t? client suppl?mentaire (m?me si aucun n?est affich? dans le diagramme). Lors du retour de la m?thode ? Async ?, l?appel reprend l?ex?cution sur le thread et le compl?ment peut acc?der aux donn?es, les exploiter et afficher le r?sultat. Le m?me motif d?ex?cution asynchrone est employ? en cas d?utilisation des applications h?tes de client riche Office, telles que Word 2013 ou Excel 2013.

*Figure 1. Flux d?ex?cution de programmation asynchrone*

![Flux d?ex?cution de thread de programmation asynchrone](../images/office15-app-async-prog-fig01.png)

La prise en charge de cette conception asynchrone dans les clients riches et les clients web fait partie des objectifs de conception ? ?criture unique-ex?cution multiplateforme ? du mod?le de d?veloppement des Compl?ments Office. Par exemple, vous pouvez cr?er un compl?ment de contenu ou du volet de t?ches avec une seule base de code qui sera ex?cut?e sur Excel 2013 et Excel Online.

## <a name="writing-the-callback-function-for-an-async-method"></a>?criture de la fonction de rappel pour une m?thode ? Async ?


La fonction de rappel que vous transmettez en tant qu?argument _callback_ ? une m?thode ? Async ? doit d?clarer un seul param?tre que le runtime de compl?ment va utiliser pour permettre l?acc?s ? un objet [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) lorsque la fonction de rappel sera ex?cut?e. Vous pouvez ?crire :


- une fonction anonyme devant ?tre ?crite et pass?e directement en ligne avec l?appel ? la m?thode ? Async ? en tant que param?tre  _callback_ de la m?thode ? Async ? ;
    
- une fonction nomm?e, en passant le nom de cette fonction en tant que param?tre  _callback_ de la m?thode ? Async ?.
    
Une fonction anonyme est utile si vous envisagez de n?utiliser son code qu?une fois : comme elle n?a pas de nom, vous ne pouvez pas y faire r?f?rence dans une autre partie du code. Une fonction nomm?e est utile si vous voulez r?utiliser la fonction de rappel pour plusieurs m?thodes ? Async ?.


### <a name="writing-an-anonymous-callback-function"></a>?criture d?une fonction de rappel anonyme

La fonction de rappel anonyme suivante d?clare un seul param?tre nomm? `result` qui r?cup?re les donn?es ? partir de la propri?t? [AsyncResult.value](https://dev.office.com/reference/add-ins/shared/asyncresult.status) lorsque le rappel est renvoy?.


```js
function (result) {
        write('Selected data: ' + result.value);
}
```

L?exemple suivant montre comment passer cette fonction de rappel anonyme dans le contexte d?un appel complet de m?thode ? Async ? ? la m?thode  **Document.getSelectedDataAsync**.


- Le premier argument  _coercionType_,  `Office.CoercionType.Text`, sp?cifie le retour des donn?es s?lectionn?es en tant que cha?ne de texte.
    
- Le deuxi?me argument  _callback_ est la fonction anonyme pass?e en ligne ? la m?thode. Lors de l?ex?cution de la fonction, elle utilise le param?tre _result_ pour acc?der ? la propri?t? **value** de l?objet **AsyncResult** afin d?afficher les donn?es s?lectionn?es par l?utilisateur dans le document.
    



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

Vous pouvez ?galement utiliser le param?tre de votre fonction de rappel pour acc?der aux autres propri?t?s de l?objet **AsyncResult**. Utilisez la propri?t? [AsyncResult.status](https://dev.office.com/reference/add-ins/shared/asyncresult.error) pour d?terminer si l?appel a r?ussi ou ?chou?. En cas d??chec, vous pouvez utiliser la propri?t? [AsyncResult.error](https://dev.office.com/reference/add-ins/shared/asyncresult.context) pour acc?der ? un objet [Error](https://dev.office.com/reference/add-ins/shared/error) et obtenir des informations sur l?erreur.

Pour plus d?informations sur l?utilisation de la m?thode  **getSelectedDataAsync**, voir [Lecture et ?criture de donn?es dans la s?lection active d?un document ou d?une feuille de calcul](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md). 


### <a name="writing-a-named-callback-function"></a>?criture d?une fonction de rappel nomm?e

Vous pouvez ?galement ?crire une fonction nomm?e et passer son nom au param?tre  _callback_ d?une m?thode ? Async ?. Par exemple, l?exemple pr?c?dent peut ?tre r??crit pour passer une fonction nomm?e `writeDataCallback` en tant que param?tre _callback_ comme suit.


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


## <a name="differences-in-whats-returned-to-the-asyncresultvalue-property"></a>Diff?rences dans les ?l?ments retourn?s ? la propri?t? AsyncResult.value


Les propri?t?s  **asyncContext**,  **status** et **error** de l?objet **AsyncResult** retournent le m?me type d?informations ? la fonction de rappel pass?e ? toutes les m?thodes ? Async ?. Cependant, les ?l?ments retourn?s ? la propri?t? **AsyncResult.value** varient selon la fonctionnalit? de la m?thode ? Async ?.

Par exemple, les m?thodes **addHandlerAsync** (des objets [Binding](https://dev.office.com/reference/add-ins/shared/binding), [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart), [Document](https://dev.office.com/reference/add-ins/shared/document), [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) et [Settings](https://dev.office.com/reference/add-ins/shared/settings)) sont utilis?es pour ajouter des fonctions de gestionnaire d??v?nements aux ?l?ments repr?sent?s par ces objets. Vous pouvez acc?der ? la propri?t? **AsyncResult.value** ? partir de la fonction de rappel que vous transmettez aux m?thodes **addHandlerAsync**, mais comme vous n?acc?dez ? aucune donn?e ni ? aucun objet lorsque vous ajoutez un gestionnaire d??v?nements, la propri?t? **value** renvoie toujours **undefined** si vous tentez d?y acc?der.

En revanche, si vous appelez la m?thode  **Document.getSelectedDataAsync**, celle-ci renvoie les donn?es que l?utilisateur a s?lectionn?es dans le document ? la propri?t?  **AsyncResult.value** dans le rappel. Ou alors, si vous appelez la m?thode [Bindings.getAllAsync](https://dev.office.com/reference/add-ins/shared/bindings.getallasync), celle-ci renvoie un tableau de tous les objets  **Binding** du document. Enfin, si vous appelez la m?thode [Bindings.getByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync), celle-ci renvoie un seul objet  **Binding**.

Pour obtenir une description des ?l?ments renvoy?s ? la propri?t? **AsyncResult.value** pour une m?thode ? Async ?, voir la section relative ? la valeur de rappel dans la rubrique de r?f?rence de cette m?thode. Pour obtenir un r?sum? de tous les objets qui fournissent des m?thodes ? Async ?, voir le tableau situ? au bas de la rubrique relative ? l?objet [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult).


## <a name="asynchronous-programming-patterns"></a>Mod?les de programmation asynchrone


L?interface API JavaScript pour Office prend en charge deux types de mod?les de programmation asynchrone :


- Utilisation des rappels imbriqu?s
    
- Utilisation du mod?le des promesses
    
La programmation asynchrone ? l?aide des fonctions de rappel n?cessite que vous imbriquiez fr?quemment le r?sultat retourn? d?un rappel au sein d?au moins deux rappels. Pour ce faire, vous pouvez utiliser les rappels imbriqu?s de toutes les m?thodes ? Async ? de l?API.

L?utilisation des rappels imbriqu?s est un mod?le de programmation familier pour la plupart des d?veloppeurs JavaScript, mais le code contenant des rappels fortement imbriqu?s peut ?tre difficile ? lire et ? comprendre. Pour offrir une solution de remplacement aux rappels imbriqu?s, l?interface API JavaScript pour Office prend ?galement en charge l?impl?mentation du mod?le des promesses. Cependant, dans la version actuelle de l?interface API JavaScript pour Office, le mod?le des promesses fonctionne uniquement avec du code destin? aux [liaisons dans les feuilles de calcul Excel et les documents Word](bind-to-regions-in-a-document-or-spreadsheet.md).

<a name="AsyncProgramming_NestedCallbacks" />
### <a name="asynchronous-programming-using-nested-callback-functions"></a>Programmation asynchrone utilisant des fonctions de rappel imbriqu?es


Vous devez fr?quemment effectuer au moins deux op?rations asynchrones pour r?aliser une t?che. Pour ce faire, vous pouvez imbriquer un appel ? Async ? dans un autre. 

L?exemple de code suivant imbrique deux appels asynchrones. 


- D?abord, la m?thode [Bindings.getByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync) est appel?e pour acc?der ? une liaison dans le document nomm? ? MyBinding ?. L?objet **AsyncResult** renvoy? au param?tre `result` de ce rappel donne acc?s ? l?objet de liaison sp?cifi? dans la propri?t? **AsyncResult.value**.
    
- Ensuite, l?objet Binding auquel vous avez acc?d? ? partir du premier param?tre `result` est utilis? pour appeler la m?thode [Binding.getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync).
    
- Enfin, le param?tre  `result2` du rappel pass? ? la m?thode **Binding.getDataAsync** est utilis? pour afficher les donn?es dans la liaison.
    



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

Ce mod?le de rappel imbriqu? de base s?applique ? toutes les m?thodes asynchrones dans l?interface API JavaScript pour Office.

Les sections suivantes montrent comment utiliser des fonctions anonymes ou nomm?es pour des rappels imbriqu?s dans des m?thodes asynchrones.


#### <a name="using-anonymous-functions-for-nested-callbacks"></a>Utilisation des fonctions anonymes pour des rappels imbriqu?s

Dans l?exemple suivant, deux fonctions anonymes sont d?clar?es en ligne et pass?es dans les m?thodes  **getByIdAsync** et **getDataAsync** en tant que rappels imbriqu?s. Comme les fonctions sont tr?s simples, l?objet de l?impl?mentation est ?vident.


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


#### <a name="using-named-functions-for-nested-callbacks"></a>Utilisation de fonctions nomm?es pour des rappels imbriqu?s

Dans des impl?mentations complexes, il peut ?tre utile d?utiliser des fonctions nomm?es pour garantir une meilleure lisibilit?, simplicit? de gestion et possibilit? de r?utilisation du code. Dans l?exemple suivant, les deux fonctions anonymes de l?exemple dans la section pr?c?dente ont ?t? r??crites comme fonctions nomm?es  `deleteAllData` et `showResult`. Ces fonctions nomm?es sont ensuite pass?es dans les m?thodes  **getByIdAsync** et **deleteAllDataValuesAsync** comme rappels par nom.


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


### <a name="asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings"></a>Programmation asynchrone en utilisant le mod?le des promesses pour acc?der aux donn?es des liaisons


Plut?t que de transmettre une fonction de rappel et d?attendre le renvoi de la fonction pour poursuivre l?ex?cution, le motif de programmation des promesses renvoie imm?diatement un objet de promesse qui repr?sente le r?sultat souhait?. Toutefois, contrairement ? la vraie programmation synchrone, en arri?re-plan, la concr?tisation du r?sultat pr?vu est en fait diff?r?e jusqu?? ce que l?environnement d?ex?cution des compl?ments Office puisse r?aliser la demande. Un gestionnaire _onError_ est fourni pour couvrir les cas o? la demande ne peut pas ?tre remplie.

L?interface API JavaScript pour Office fournit la m?thode [Office.select](https://dev.office.com/reference/add-ins/shared/office.select) pour prendre en charge le mod?le des promesses permettant d?utiliser des objets de liaison existants. L?objet de promesse renvoy? ? la m?thode **Office.select** prend en charge uniquement les quatre m?thodes auxquelles vous pouvez acc?der directement ? partir de l?objet [Binding](https://dev.office.com/reference/add-ins/shared/binding) : [getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync), [setDataAsync](https://dev.office.com/reference/add-ins/shared/binding.setdataasync), [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/asyncresult.value) et [removeHandlerAsync](https://dev.office.com/reference/add-ins/shared/binding.removehandlerasync).

Le mod?le des promesses ? utiliser avec les liaisons se pr?sente comme suit :

 **Office.select(**_selectorExpression_,  _onError_**).**_BindingObjectAsyncMethod_

Le param?tre  _selectorExpression_ a le format `"bindings#bindingId"`, o?  _bindingId_ est le nom ( **id**) d?une liaison cr??e pr?c?demment dans le document ou la feuille de calcul (? l?aide de l?une des m?thodes ? addFrom ? de la collection  **Bindings** :  **addFromNamedItemAsync**,  **addFromPromptAsync** ou **addFromSelectionAsync**). Par exemple, l?expression de s?lecteur  `bindings#cities` sp?cifie que vous voulez acc?der ? la liaison avec le param?tre **id** 'cities'.

Le param?tre  _onError_ est une fonction de gestion des erreurs qui prend un seul param?tre de type **AsyncResult** pouvant ?tre utilis? pour acc?der ? un objet **Error** si la m?thode **select** ne permet pas d?acc?der ? la liaison sp?cifi?e. L?exemple suivant montre une fonction de gestion des erreurs de base pouvant ?tre pass?e au param?tre _onError_.




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

Remplacez l?espace r?serv? _BindingObjectAsyncMethod_ par un appel ? l?une des quatre m?thodes d?objet **Binding** prises en charge par l?objet de promesse : **getDataAsync**, **setDataAsync**, **addHandlerAsync** ou **removeHandlerAsync**. Les appels ? ces m?thodes ne prennent pas en charge les promesses suppl?mentaires. Vous devez les appeler ? l?aide du [mod?le de fonction de rappel imbriqu?e](#AsyncProgramming_NestedCallbacks).

Une fois qu?une promesse d?objet  **Binding** est concr?tis?e, elle peut ?tre r?utilis?e dans l?appel de m?thode cha?n? comme s?il s?agissait d?une liaison (le runtime de compl?ment ne retentera pas de concr?tiser la promesse de fa?on asynchrone). Si la promesse d?objet **Binding** ne peut pas ?tre concr?tis?e, le runtime de compl?ment retentera d?acc?der ? l?objet de liaison au prochain appel de l?une de ses m?thodes asynchrones.

L?exemple de code suivant utilise la m?thode **select** pour r?cup?rer une liaison avec l?**id** ? `cities` ? ? partir de la collection **Bindings**, puis appelle la m?thode [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/asyncresult.value) afin d?ajouter un gestionnaire d??v?nements pour l??v?nement [dataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) de la liaison.




```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```


> [!IMPORTANT]
> La promesse d?objet **Binding** renvoy?e par la m?thode **Office.select** fournit uniquement un acc?s aux quatre m?thodes de l?objet **Binding**. Pour acc?der ? l?un des autres membres de l?objet **Binding**, vous devez utiliser la propri?t? **Document.bindings** et la m?thode **Bindings.getByIdAsync** ou **Bindings.getAllAsync** pour r?cup?rer l?objet **Binding**. Par exemple, pour acc?der aux propri?t?s de l?objet **Binding** (propri?t? **document**, **id** ou **type**) ou pour acc?der aux propri?t?s de l?objet [MatrixBinding](https://dev.office.com/reference/add-ins/shared/binding.matrixbinding) ou [TableBinding](https://dev.office.com/reference/add-ins/shared/binding.tablebinding), vous devez utiliser la m?thode **getByIdAsync** ou **getAllAsync** pour r?cup?rer un objet **Binding**.


## <a name="passing-optional-parameters-to-asynchronous-methods"></a>Passage de param?tres facultatifs ? des m?thodes asynchrones


La syntaxe courante pour toutes les m?thodes ? Async ? suit ce mod?le :

 _AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_ `);`

Toutes les m?thodes asynchrones prennent en charge des param?tres facultatifs, qui sont pass?s sous la forme d?un objet JSON (JavaScript Object Notation) qui contient un ou plusieurs param?tres facultatifs. L?objet JSON contenant les param?tres facultatifs est une collection non ordonn?e de paires cl?-valeur o? le caract?re ? : ? s?pare la cl? de la valeur. Chaque paire dans l?objet est s?par?e par une virgule, et l?ensemble complet de paires est plac? entre accolades. La cl? est le nom du param?tre, et la valeur est la valeur ? passer pour ce param?tre.

Vous pouvez cr?er l?objet JSON qui contient les param?tres facultatifs incorpor?s, ou cr?er un objet  `options` et le passer comme param?tre _options_.


### <a name="passing-optional-parameters-inline"></a>Passage de param?tres facultatifs incorpor?s

Par exemple, la syntaxe pour appeler la m?thode [Document.setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) avec des param?tres facultatifs incorpor?s se pr?sente comme ceci :

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext:' asyncContext},callback);

```

Dans cette forme de syntaxe d?appel, les deux param?tres facultatifs,  _coercionType_ et _asyncContext_, sont d?finis comme un objet incorpor? mis entre accolades.

L?exemple suivant montre comment appeler la m?thode **Document.setSelectedDataAsync** en sp?cifiant des param?tres facultatifs incorpor?s.


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
> Vous pouvez sp?cifier des param?tres facultatifs dans l?objet JSON dans n?importe quel ordre dans la mesure o? leurs noms sont correctement sp?cifi?s.


### <a name="passing-optional-parameters-in-an-options-object"></a>Passage de param?tres facultatifs dans un objet options

Vous pouvez ?galement cr?er un objet nomm?  `options` qui sp?cifie les param?tres facultatifs s?par?ment de l?appel de la m?thode, puis passe l?objet `options` comme l?argument _options_.

L?exemple suivant illustre une mani?re de cr?er l?objet  `options`, o?  `parameter1` et `value1` notamment sont des espaces r?serv?s aux noms et valeurs de param?tres effectifs.




```js
var options = {
    parameter1: value1,
    parameter2: value2,
    ...
    parameterN: valueN
};

```

Ce qui ressemble ? l?exemple suivant lors de la sp?cification des param?tres [ValueFormat](https://dev.office.com/reference/add-ins/shared/valueformat-enumeration) et [FilterType](https://dev.office.com/reference/add-ins/shared/filtertype-enumeration).




```js
var options = {
    valueFormat: "unformatted",
    filterType: "all"
};
```

Voici une autre fa?on de cr?er l?objet  `options`.




```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

Ce qui ressemble ? l?exemple suivant lors de la sp?cification des param?tres  **ValueFormat** et **FilterType** :


```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```


> [!NOTE]
> Au moment de cr?er l?objet `options` en employant l?une ou l?autre de ces m?thodes, vous pouvez sp?cifier des param?tres facultatifs dans n?importe quel ordre du moment o? leurs noms sont sp?cifi?s correctement.

L?exemple suivant illustre comment appeler la m?thode **Document.setSelectedDataAsync** en sp?cifiant des param?tres facultatifs dans un objet `options`.




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


Dans les deux exemples de param?tres facultatifs, le param?tre _callback_ est sp?cifi? comme le dernier param?tre (? la suite des param?tres facultatifs incorpor?s, ou de l?objet de l?argument _options_). Vous pouvez ?galement sp?cifier le param?tre _callback_ ? l?int?rieur de l?objet JSON incorpor?, ou dans l?objet `options`. Cependant, vous ne pouvez passer le param?tre _callback_ qu?? un seul endroit : soit dans l?objet _options_ (incorpor? ou cr?? en externe), soit comme dernier param?tre, mais pas les deux.


## <a name="see-also"></a>Voir aussi

- [Pr?sentation de l?API JavaScript pour Office](understanding-the-javascript-api-for-office.md) 
- [Interface API JavaScript pour Office](https://dev.office.com/reference/add-ins/javascript-api-for-office)
     
