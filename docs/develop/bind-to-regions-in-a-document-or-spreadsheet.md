---
title: Lier des r?gions dans un document ou une feuille de calcul
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: bd26aa12e5d6da145fb6a2a89daf937cf6e88f04
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="bind-to-regions-in-a-document-or-spreadsheet"></a>Lier des r?gions dans un document ou une feuille de calcul

L?acc?s aux donn?es bas?es sur une liaison permet aux compl?ments de contenu et du volet Office d?acc?der de fa?on coh?rente ? une zone particuli?re d?un document ou d?une feuille de calcul au moyen d?un identificateur. Le compl?ment doit d?abord ?tablir la liaison en appelant l?une des m?thodes qui associent une partie du document ? un identificateur unique : [addFromPromptAsync], [addFromSelectionAsync] ou [addFromNamedItemAsync]. Une fois la liaison ?tablie, le compl?ment peut utiliser l?identificateur fourni pour acc?der aux donn?es contenues dans la zone associ?e du document ou de la feuille de calcul. La cr?ation de liaisons apporte la valeur ajout?e suivante ? votre compl?ment :


- Elle permet l?acc?s aux structures de donn?es communes sur les applications Office prises en charge, telles que : tableaux, plages ou texte (s?rie contigu? de caract?res).
    
- Elle permet les op?rations de lecture/?criture sans exiger que l?utilisateur effectue une s?lection.
    
- Elle ?tablit une relation entre le compl?ment et les donn?es du document. Les liaisons persistent dans le document et sont accessibles par la suite.
    
L??tablissement d?une liaison vous permet ?galement de vous abonner aux donn?es et aux ?v?nements de changement de s?lection qui sont concern?s par cette r?gion particuli?re du document ou de la feuille de calcul. Cela signifie que le compl?ment est seulement notifi? des changements qui surviennent dans la r?gion d?limit?e, par opposition aux changements g?n?raux affectant l?ensemble du document ou de la feuille de calcul.

L?objet [Bindings] expose une m?thode [getAllAsync] qui donne acc?s ? toutes les liaisons ?tablies dans le document ou la feuille de calcul. Une liaison individuelle est accessible par son ID ? l?aide de la m?thode [Bindings.getBindingByIdAsync] ou [Office.select]. Vous pouvez ?tablir de nouvelles liaisons et supprimer des liaisons existantes en utilisant l?une des m?thodes suivantes de l?objet [Bindings] : [addFromSelectionAsync], [addFromPromptAsync], [addFromNamedItemAsync] ou [releaseByIdAsync].


## <a name="binding-types"></a>Types de liaison

Vous sp?cifiez [trois types de liaisons diff?rents][Office.BindingType] avec le param?tre _bindingType_ lorsque vous cr?ez une liaison avec les m?thodes [addFromSelectionAsync], [addFromPromptAsync] ou [addFromNamedItemAsync] :

1. **[Liaison de texte][TextBinding]** - ?tablit une liaison ? une zone du document qui est repr?sent?e en tant que texte.

    Dans Word, la plupart des s?lections contigu?s sont valides, tandis que dans Excel, seules les s?lections de cellules uniques peuvent ?tre la cible d?une liaison de texte. Dans Excel, seul le texte brut est pris en charge. Dans Word, trois formats sont pris en charge : texte brut, HTML et Open XML pour Office.

2. **[Liaison de matrice][MatrixBinding]** - ?tablit une liaison ? une zone d?un document qui contient des donn?es tabulaires sans en-t?te. Les donn?es dans une liaison de matrice sont ?crites ou lues comme un **tableau** bidimensionnel, ce qui est impl?ment? sous la forme d?un tableau de tableaux dans JavaScript. Par exemple, deux lignes d?une valeur de **cha?ne** dans deux colonnes peuvent ?tre ?crites ou lues comme ` [['a', 'b'], ['c', 'd']]`, et une colonne unique de trois lignes peut ?tre ?crite ou lue comme `[['a'], ['b'], ['c']]`.

    Dans Excel, toute s?lection contigu? de cellules peut ?tre utilis?e pour ?tablir une liaison de matrice. Dans Word, seuls les tableaux prennent en charge la liaison de matrice.

3. **[Liaison de tableau][TableBinding]** - ?tablit une liaison ? une zone d?un document qui contient un tableau avec des en-t?tes. Les donn?es dans une liaison de tableau sont ?crites ou lues comme un objet [TableData](https://dev.office.com/reference/add-ins/shared/tabledata). L?objet `TableData` expose les donn?es via les propri?t?s `headers` et `rows`.

    Tout tableau Excel ou Word peut ?tre la base d?une liaison de tableau. Une fois que vous ?tablissez une liaison de tableau, chaque nouvelle ligne ou colonne qu?un utilisateur ajoute au tableau est automatiquement incluse dans la liaison.

Apr?s la cr?ation d?une liaison ? l?aide de l?une des trois m?thodes ? addFrom ? de l?objet `Bindings`, vous pouvez travailler avec les donn?es et les propri?t?s de la liaison en utilisant les m?thodes de l?objet correspondant : [MatrixBinding], [TableBinding] ou [TextBinding]. Ces trois objets h?ritent des m?thodes [getDataAsync] et [setDataAsync] de l?objet `Binding` qui vous permettent d?interagir avec les donn?es li?es.

> [!NOTE]
> **Quand devez-vous utiliser une liaison de matrice ou une liaison de tableau ?** Lorsque les donn?es tabulaires avec lesquelles vous travaillez contiennent une ligne de total, vous devez utiliser une liaison de matrice si le script de votre compl?ment doit acc?der aux valeurs figurant dans la ligne de total ou d?tecter que la s?lection de l?utilisateur figure dans la ligne de total. Si vous ?tablissez une liaison de tableau pour des donn?es tabulaires qui contiennent une ligne de total, la propri?t? [TableBinding.rowCount] et les propri?t?s `rowCount` et `startRow` de l?objet [BindingSelectionChangedEventArgs] dans les gestionnaires d??v?nements ne refl?teront pas la ligne de total dans leurs valeurs. Pour contourner cette limitation, vous devez ?tablir une liaison de matrice pour travailler avec la ligne de total.

## <a name="add-a-binding-to-the-users-current-selection"></a>Ajout d?une liaison ? la s?lection actuelle de l?utilisateur

L?exemple suivant montre comment ajouter une liaison de texte nomm?e `myBinding` ? la s?lection actuelle dans un document ? l?aide de la m?thode [addFromSelectionAsync].


```js
Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Dans cet exemple, le type de liaison sp?cifi? est ? Text ?. Cela signifie qu?un objet [TextBinding] sera cr?? pour la s?lection. Diff?rents types de liaison exposent diff?rentes donn?es et op?rations. [Office.BindingType] est une ?num?ration des valeurs de types de liaison disponibles.

Le deuxi?me param?tre facultatif est un objet qui sp?cifie l?ID de la nouvelle liaison cr??e. Si un ID n?est pas sp?cifi?, un ID est g?n?r? automatiquement.

La fonction anonyme qui est pass?e dans la fonction comme param?tre final _callback_ est ex?cut?e lorsque la cr?ation de la liaison est termin?e. La fonction est appel?e avec un seul param?tre, `asyncResult`, ce qui donne acc?s ? un objet [AsyncResult] qui fournit l??tat de l?appel. La propri?t? `AsyncResult.value` contient une r?f?rence ? un objet [Binding] du type sp?cifi? pour la liaison cr??e r?cemment. Vous pouvez utiliser cet objet [Binding] pour obtenir et d?finir les donn?es.

## <a name="add-a-binding-from-a-prompt"></a>Ajout d?une liaison ? partir d?une invite

L?exemple suivant indique comment ajouter une liaison de texte appel?e `myBinding` ? l?aide de la m?thode [addFromPromptAsync]. Cette m?thode permet ? l?utilisateur de sp?cifier la plage pour la liaison ? l?aide de l?invite de s?lection de plage int?gr?e.


```js
function bindFromPrompt() {
    Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        } else {
            write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Dans cet exemple, le type de liaison sp?cifi? est ? Text ?. Cela signifie qu?un objet [TextBinding] sera cr?? pour la s?lection que l?utilisateur sp?cifie dans l?invite.

Le deuxi?me param?tre est un objet qui contient l?ID de la nouvelle liaison cr??e. Si un ID n?est pas sp?cifi?, un ID est g?n?r? automatiquement.

La fonction anonyme transmise dans la fonction comme troisi?me param?tre _callback_ est ex?cut?e lorsque la cr?ation de la liaison est termin?e. Lorsque la fonction de rappel s?ex?cute, l?objet [AsyncResult] contient le statut de l?appel et la nouvelle liaison.

La figure 1 montre l?invite de s?lection de plage int?gr?e dans Excel.


*Figure 1. Interface utilisateur de s?lection de donn?es dans Excel*

![Interface utilisateur de s?lection de donn?es dans Excel](../images/agave-api-overview-excel-selection-ui.png)


## <a name="add-a-binding-to-a-named-item"></a>Ajout d?une liaison ? un ?l?ment nomm?


L?exemple suivant montre comment ajouter une liaison de matrice ? l??l?ment nomm? `myRange` existant en utilisant la m?thode [addFromNamedItemAsync], et d?finit le param?tre `id` de la liaison sur ? myMatrix ?.


```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("myRange", "matrix", {id:'myMatrix'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

**Pour Excel**, le param?tre `itemName` de la m?thode [addFromNamedItemAsync] peut se r?f?rer ? une plage nomm?e existante, une plage sp?cifi?e avec le style de r?f?rence `A1` `("A1:A3")` ou un tableau. Par d?faut, l?ajout d?un tableau dans Excel entra?ne l?affectation du nom ? Tableau1 ? pour le premier tableau que vous ajoutez, ? Tableau2 ? pour le deuxi?me tableau que vous ajoutez, et ainsi de suite. Pour affecter un nom significatif ? un tableau dans l?interface utilisateur d?Excel, servez-vous de la propri?t? **Table Name** sous l?onglet **Outils de tableau | Conception** du ruban.


> [!NOTE]
> Dans Excel, lors de la sp?cification d?un tableau comme ?l?ment nomm?, vous devez enti?rement qualifier le nom pour inclure le nom de la feuille de calcul dans le nom du tableau dans ce format :  `"Sheet1!Table1"`

L?exemple suivant cr?e une liaison dans Excel aux trois premi?res cellules de la colonne A (`"A1:A3"`), attribue l?id`"MyCities"`, puis ?crit trois noms de ville dans cette liaison.


```js
 function bindingFromA1Range() {
    Office.context.document.bindings.addFromNamedItemAsync("A1:A3", "matrix", {id: "MyCities" },
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                write('Error: ' + asyncResult.error.message);
            }
            else {
                // Write data to the new binding.
                Office.select("bindings#MyCities").setDataAsync([['Berlin'], ['Munich'], ['Duisburg']], { coercionType: "matrix" },
                    function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            write('Error: ' + asyncResult.error.message);
                        }
                    });
            }
        });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

**Pour Word**, le param?tre `itemName` de la m?thode [addFromNamedItemAsync] fait r?f?rence ? la propri?t? `Title` d?un contr?le de contenu `Rich Text`. (Vous ne pouvez r?aliser de liaison avec des contr?les de contenu diff?rents du contr?le de contenu `Rich Text`.)

Par d?faut, un contr?le de contenu ne comporte aucune valeur affect?e `Title*`. Pour attribuer un nom significatif dans l?interface utilisateur de Word, apr?s avoir ins?r? un contr?le de contenu de **texte enrichi** ? partir du groupe **Contr?les** sous l?onglet **D?veloppeur** du ruban, utilisez la commande **Propri?t?s** dans le groupe **Contr?les** pour afficher la bo?te de dialogue **Propri?t?s du contr?le de contenu**. D?finissez la propri?t? **Title** du contr?le de contenu sur le nom auquel vous souhaitez faire r?f?rence ? partir de votre code.

L?exemple suivant cr?e une liaison de texte dans Word vers un contr?le de contenu de texte enrichi nomm?  `"FirstName"`, attribue l? **id**`"firstName"`, puis affiche cette information.


```js
function bindContentControl() {
    Office.context.document.bindings.addFromNamedItemAsync('FirstName', 
        Office.BindingType.Text, {id:'firstName'},
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                write('Control bound. Binding.id: '
                    + result.value.id + ' Binding.type: ' + result.value.type);
            } else {
                write('Error:', result.error.message);
            }
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

## <a name="get-all-bindings"></a>Obtention de toutes les liaisons


L?exemple suivant montre comment obtenir toutes les liaisons dans un document en utilisant la m?thode Bindings.[getAllAsync].


```js
Office.context.document.bindings.getAllAsync(function (asyncResult) {
    var bindingString = '';
    for (var i in asyncResult.value) {
        bindingString += asyncResult.value[i].id + '\n';
    }
    write('Existing bindings: ' + bindingString);
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

La fonction anonyme qui est pass?e dans la fonction comme param?tre `callback` est ex?cut?e lorsque l?op?ration est termin?e. La fonction est appel?e avec un seul param?tre, `asyncResult`, qui contient un tableau des liaisons dans le document. Le tableau est r?p?t? pour g?n?rer une cha?ne qui contient les ID des liaisons. La cha?ne est ensuite affich?e dans une bo?te de message.


## <a name="get-a-binding-by-id-using-the-getbyidasync-method-of-the-bindings-object"></a>Obtention d?une liaison par ID en utilisant la m?thode getByIdAsync de l?objet Bindings


L?exemple suivant indique comment utiliser la m?thode [getByIdAsync] pour obtenir une liaison dans un document en sp?cifiant son ID. Cet exemple suppose qu?une liaison nomm?e `'myBinding'` a ?t? ajout?e au document ? l?aide des m?thodes d?crites plus haut dans cette rubrique.


```js
Office.context.document.bindings.getByIdAsync('myBinding', function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } 
    else {
        write('Retrieved binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Dans l?exemple, le premier param?tre `id` est l?ID de la liaison ? r?cup?rer.

La fonction anonyme qui est pass?e dans la fonction comme second param?tre  _callback_ est ex?cut?e lorsque l?op?ration est termin?e. La fonction est appel?e avec un seul param?tre, _asyncResult_, qui contient le statut de l?appel et la liaison avec l?ID ? myBinding ?.


## <a name="get-a-binding-by-id-using-the-select-method-of-the-office-object"></a>Obtention d?une liaison par ID en utilisant la m?thode Select de l?objet Office


L?exemple suivant montre comment utiliser la m?thode [Office.select] pour obtenir une promesse d?objet [Binding] dans un document en sp?cifiant son ID dans une cha?ne de s?lecteur. Il appelle ensuite la m?thode [Binding.getDataAsync] pour obtenir des donn?es ? partir de la liaison sp?cifi?e. Cet exemple suppose qu?une liaison nomm?e `'myBinding'` a ?t? ajout?e au document ? l?aide des m?thodes d?crites plus haut dans cette rubrique.


```js
Office.select("bindings#myBinding", function onError(){}).getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


> [!NOTE]
> Si la promesse de la m?thode `select` renvoie un objet [Binding], cet objet expose uniquement les quatre m?thodes suivantes de l?objet : [getDataAsync], [setDataAsync], [addHandlerAsync] et [removeHandlerAsync]. Si la promesse ne peut pas renvoyer un objet Binding, le rappel `onError` peut ?tre utilis? pour acc?der ? un objet [asyncResult].error afin d?obtenir plus d?informations. Si vous devez appeler un membre de l?objet Binding autre que les quatre m?thodes expos?es par la promesse d?objet Binding renvoy?e par la m?thode `select`, utilisez plut?t la m?thode [getByIdAsync] en employant la propri?t? [Document.bindings] et la m?thode [Bindings.getByIdAsync] pour r?cup?rer l?objet Binding**.

## <a name="release-a-binding-by-id"></a>Publication d?une liaison par ID


L?exemple suivant montre comment utiliser la m?thode [releaseByIdAsync] pour publier une liaison dans un document en sp?cifiant son ID.

```js
Office.context.document.bindings.releaseByIdAsync('myBinding', function (asyncResult) {
    write('Released myBinding!');
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Dans l?exemple, le premier param?tre `id` est l?ID de la liaison ? publier.

La fonction anonyme qui est pass?e dans la fonction comme le deuxi?me param?tre est un rappel qui est ex?cut? lorsque l?op?ration est termin?e. La fonction est appel?e avec un seul param?tre,  [asyncResult], qui contient le statut de l?appel.


## <a name="read-data-from-a-binding"></a>Lecture de donn?es ? partir d?une liaison


L?exemple suivant montre comment utiliser la m?thode [getDataAsync] pour obtenir des donn?es ? partir d?une liaison existante.


```js
myBinding.getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 `myBinding` est une variable qui contient une liaison de texte existante dans le document. Vous pouvez ?galement utiliser [Office.select] pour acc?der ? la liaison avec son identifiant et commencer ? appeler la m?thode [getDataAsync] de la mani?re suivante : 

```js 
Office.select("bindings#myBindingID").getDataAsync
```


La fonction anonyme qui est pass?e dans la fonction est un rappel qui est ex?cut? lorsque l?op?ration est termin?e. La propri?t? [AsyncResult].value contient les donn?es dans `myBinding`. Le type de valeur d?pend du type de liaison. La liaison dans cet exemple est une liaison de texte. Par cons?quent, la valeur contiendra une cha?ne. Pour obtenir des exemples suppl?mentaires concernant l?utilisation des liaisons de matrice et de tableau, consultez la rubrique sur la m?thode [getDataAsync].


## <a name="write-data-to-a-binding"></a>?criture de donn?es dans une liaison

L?exemple suivant montre comment utiliser la m?thode [setDataAsync] pour d?finir des donn?es dans une liaison existante.

```js
myBinding.setDataAsync('Hello World!', function (asyncResult) { });
```

 `myBinding` est une variable qui contient une liaison de texte existante dans le document.

Dans l?exemple, le premier param?tre est la valeur ? d?finir sur `myBinding`. Comme il s?agit d?une liaison de texte, la valeur est de type `string`. Diff?rents types de liaisons acceptent divers types de donn?es.

La fonction anonyme qui est pass?e dans la fonction est un rappel qui est ex?cut? lorsque l?op?ration est termin?e. La fonction est appel?e avec un seul param?tre, `asyncResult`, qui contient l??tat du r?sultat.

> [!NOTE]
> Depuis la publication d?Excel 2013 SP1 et de la version correspondante d?Excel Online, vous pouvez d?sormais [d?finir la mise en forme lors de l??criture et de la mise ? jour des donn?es dans des tableaux li?s](../excel/excel-add-ins-tables.md).


## <a name="detect-changes-to-data-or-the-selection-in-a-binding"></a>D?tection des modifications apport?es aux donn?es ou ? la section dans une liaison


L?exemple suivant montre comment lier un gestionnaire d??v?nements ? l??v?nement [DataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) d?une liaison ayant l?ID ? MyBinding ?.


```js
function addHandler() {
Office.select("bindings#MyBinding").addHandlerAsync(
    Office.EventType.BindingDataChanged, dataChanged);
}
function dataChanged(eventArgs) {
    write('Bound data changed in binding: ' + eventArgs.binding.id);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

est une variable qui contient une liaison de texte existante dans le document.`myBinding`

Le premier param?tre `eventType` de la m?thode [addHandlerAsync] sp?cifie le nom de l??v?nement auquel s?abonner. [Office.EventType] est une ?num?ration des valeurs de types d??v?nement disponibles. `Office.EventType.BindingDataChanged evaluates to the string `"bindingDataChanged"`.

La fonction  `dataChanged` qui est pass?e dans la fonction comme deuxi?me param?tre _handler_ est un gestionnaire d??v?nements qui est ex?cut? lorsque les donn?es dans la liaison sont modifi?es. La fonction est appel?e avec un seul param?tre, _eventArgs_, qui contient une r?f?rence ? la liaison. Cette liaison peut ?tre utilis?e pour r?cup?rer les donn?es mises ? jour.

De m?me, vous pouvez d?tecter lorsqu?un utilisateur modifie la s?lection dans une liaison en ajoutant un gestionnaire d??v?nements ? l??v?nement [SelectionChanged] d?une liaison. Pour ce faire, sp?cifiez le param?tre `eventType` de la m?thode [addHandlerAsync] comme `Office.EventType.BindingSelectionChanged` ou `"bindingSelectionChanged"`.

Vous pouvez ajouter plusieurs gestionnaires d??v?nements pour un ?v?nement donn? en appelant ? nouveau la m?thode [addHandlerAsync] et en transmettant une fonction de gestionnaire d??v?nements suppl?mentaire pour le param?tre `handler`. Cela fonctionnera correctement tant que le nom de chaque fonction de gestionnaire d??v?nements est unique.


### <a name="remove-an-event-handler"></a>Suppression d?un gestionnaire d??v?nements


Pour supprimer un gestionnaire d??v?nements pour un ?v?nement, appelez la m?thode [removeHandlerAsync] en transmettant le type d??v?nement en tant que premier param?tre _eventType_, puis le nom de la fonction de gestionnaire d??v?nements ? supprimer comme deuxi?me param?tre _handler_. Par exemple, la fonction suivante supprimera la fonction de gestionnaire d??v?nements `dataChanged` ajout?e dans l?exemple de la section pr?c?dente.


```js
function removeEventHandlerFromBinding() {
    Office.select("bindings#MyBinding").removeHandlerAsync(
        Office.EventType.BindingDataChanged, {handler:dataChanged});
}
```


> [!IMPORTANT]
> Si le param?tre facultatif _handler_ est omis lors de l?appel ? la m?thode [removeHandlerAsync], tous les gestionnaires d??v?nements du param?tre `eventType` sp?cifi? seront supprim?s.


## <a name="see-also"></a>Voir aussi

- [Pr?sentation de l?API JavaScript pour Office](understanding-the-javascript-api-for-office.md) 
- [Programmation asynchrone dans des compl?ments Office](asynchronous-programming-in-office-add-ins.md)
- [Lecture et ?criture de donn?es dans la s?lection active d?un document ou d?une feuille de calcul](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
[Binding]:               https://dev.office.com/reference/add-ins/shared/binding
[MatrixBinding]:         https://dev.office.com/reference/add-ins/shared/binding.matrixbinding
[TableBinding]:          https://dev.office.com/reference/add-ins/shared/binding.tablebinding
[TextBinding]:           https://dev.office.com/reference/add-ins/shared/binding.textbinding
[getDataAsync]:          https://dev.office.com/reference/add-ins/shared/binding.getdataasync
[setDataAsync]:          https://dev.office.com/reference/add-ins/shared/binding.setdataasync
[SelectionChanged]:      https://dev.office.com/reference/add-ins/shared/binding.bindingselectionchangedevent
[addHandlerAsync]:       https://dev.office.com/reference/add-ins/shared/binding.addhandlerasync
[removeHandlerAsync]:    https://dev.office.com/reference/add-ins/shared/binding.removehandlerasync

[Bindings]:              https://dev.office.com/reference/add-ins/shared/bindings.bindings
[getByIdAsync]:          https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync 
[getAllAsync]:           https://dev.office.com/reference/add-ins/shared/bindings.getallasync
[addFromNamedItemAsync]: https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync
[addFromSelectionAsync]: https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync
[addFromPromptAsync]:    https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync
[releaseByIdAsync]:      https://dev.office.com/reference/add-ins/shared/bindings.releasebyidasync

[AsyncResult]:          https://dev.office.com/reference/add-ins/shared/asyncresult
[Office.BindingType]:   https://dev.office.com/reference/add-ins/shared/bindingtype-enumeration
[Office.select]:        https://dev.office.com/reference/add-ins/shared/office.select 
[Office.EventType]:     https://dev.office.com/reference/add-ins/shared/eventtype-enumeration 
[Document.bindings]:    https://dev.office.com/reference/add-ins/shared/document.bindings


[TableBinding.rowCount]: https://dev.office.com/reference/add-ins/shared/binding.tablebinding.rowcount
[BindingSelectionChangedEventArgs]: https://dev.office.com/reference/add-ins/shared/binding.bindingselectionchangedeventargs
