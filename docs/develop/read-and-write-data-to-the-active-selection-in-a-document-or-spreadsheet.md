---
title: Lire et écrire des données dans la sélection active d’un document ou d’une feuille de calcul
description: ''
ms.date: 12/04/2017
---


# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>Lire et écrire des données dans la sélection active d’un document ou d’une feuille de calcul

L’objet [Document](https://dev.office.com/reference/add-ins/shared/document) expose des méthodes qui vous permettent de lire et d’écrire dans la sélection active de l’utilisateur dans un document ou une feuille de calcul. Pour cela, l’objet **Document** fournit les méthodes **getSelectedDataAsync** et **setSelectedDataAsync**. Cette rubrique explique comment lire, écrire et créer des gestionnaires d’événements pour détecter les changements intervenant dans la sélection de l’utilisateur.

La méthode **getSelectedDataAsync** ne fonctionne que sur la sélection active de l’utilisateur. Si vous devez conserver la sélection dans le document, afin que la même sélection soit disponible en lecture et en écriture dans les sessions exécutant votre complément, vous devez ajouter une liaison à l’aide de la méthode [Bindings.addFromSelectionAsync](https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync) (ou créer une liaison à l’aide de l’une des autres méthodes « addFrom » de l’objet [Bindings](https://dev.office.com/reference/add-ins/shared/bindings.bindings)). Pour plus d’informations sur la création d’une liaison vers une zone d’un document et sur la lecture et l’écriture dans une liaison, voir [Liaisons de régions dans un document ou une feuille de calcul](bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="read-selected-data"></a>Lecture de données sélectionnées


L’exemple suivant montre comment obtenir les données d’une sélection dans un document en utilisant la méthode [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync).


```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    }
    else {
        write('Selected data: ' + asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Dans cet exemple, le premier paramètre _coercionType_ est spécifié comme **Office.CoercionType.Text** (vous pouvez également spécifier ce paramètre en utilisant la chaîne littérale `"text"`). Cela signifie que la propriété [value](https://dev.office.com/reference/add-ins/shared/asyncresult.status) de l’objet [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) qui est disponible à partir du paramètre _asyncResult_ dans la fonction de rappel renverra une **string** qui contient le texte sélectionné dans le document. La spécification de différents types de forçage de type produit des valeurs différentes. [Office.CoercionType](https://dev.office.com/reference/add-ins/shared/coerciontype-enumeration) est une énumération des valeurs de types de forçage de type disponibles. **Office.CoercionType.Text** prend la valeur de la chaîne « text ».


> [!TIP]
> **Quand devez-vous utiliser la matrice ou le paramètre coercionType de tableau pour accéder aux données ?** Si les données tabulaires sélectionnées doivent croître de façon dynamique lors de l’ajout de lignes et de colonnes, et que vous devez travailler avec des en-têtes de tableaux, vous devez utiliser le type de données de tableau (en spécifiant le paramètre _coercionType_ de la méthode **getSelectedDataAsync** en tant que `"table"` ou **Office.CoercionType.Table**). L’ajout de lignes et de colonnes au sein de la structure de données est pris en charge dans les données de tableau et de matrice, mais l’ajout de lignes et de colonnes à la fin est pris en charge uniquement pour les données de tableau. Si vous ne prévoyez pas d’ajouter des lignes et des colonnes, et que vos données ne nécessitent pas la fonctionnalité d’en-tête, vous devez utiliser le type de données de matrice (en spécifiant le paramètre _coercionType_ de la méthode **getSelecteDataAsync** en tant que `"matrix"` ou **Office.CoercionType.Matrix**), qui fournit un modèle plus simple d’interaction avec les données.

La fonction anonyme qui est transmise dans la fonction comme deuxième paramètre _callback_ est exécutée lorsque l’opération **getSelectedDataAsync** est terminée. La fonction est appelée avec un seul paramètre, _asyncResult_, qui contient le résultat et l’état de l’appel. Si l’appel échoue, la propriété [error](https://dev.office.com/reference/add-ins/shared/asyncresult.context) de l’objet **AsyncResult** donne accès à l’objet [Error](https://dev.office.com/reference/add-ins/shared/error). Vous pouvez vérifier la valeur des propriétés [Error.name](https://dev.office.com/reference/add-ins/shared/error.name) et [Error.message](https://dev.office.com/reference/add-ins/shared/error.message) pour déterminer les raisons de l’échec de l’opération. Sinon, le texte sélectionné dans le document s’affiche.

La propriété [AsyncResult.status](https://dev.office.com/reference/add-ins/shared/asyncresult.error) est utilisée dans l’instruction **if** pour tester la réussite de l’appel. [Office.AsyncResultStatus](https://dev.office.com/reference/add-ins/shared/asyncresultstatus-enumeration) est une énumération des valeurs de propriété **AsyncResult.status** disponibles. **Office.AsyncResultStatus.Failed** prend la valeur de la chaîne « failed » (et, de nouveau, peut également être spécifié comme chaîne littérale).


## <a name="write-data-to-the-selection"></a>Écriture de données dans la sélection


L’exemple suivant montre comment définir la sélection pour afficher « Hello World! ».


```js
Office.context.document.setSelectedDataAsync("Hello World!", function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Le passage de différents types d’objets pour le paramètre  _data_ produit différents résultats. Le résultat varie en fonction de la sélection actuelle dans le document, de l’application qui héberge votre complément, et de l’éventuel passage forcé des données dans la sélection actuelle.

La fonction anonyme transmise dans la méthode [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) comme paramètre _callback_ est exécutée quand l’appel anonyme est terminé. Lorsque vous écrivez des données dans la sélection à l’aide de la méthode **setSelectedDataAsync**, le paramètre _asyncResult_ du rappel donne uniquement accès à l’état de l’appel et à l’objet [Error](https://dev.office.com/reference/add-ins/shared/error) si l’appel échoue.

> [!NOTE]
> Depuis la publication d’Excel 2013 SP1 et de la version correspondante d’Excel Online, vous pouvez désormais [définir la mise en forme lors de l’écriture d’un tableau sur la sélection active](../excel/excel-add-ins-tables.md).


## <a name="detect-changes-in-the-selection"></a>Détection de modifications dans la sélection


L’exemple suivant montre comment détecter des modifications dans la sélection à l’aide de la méthode [Document.addHandlerAsync](https://dev.office.com/reference/add-ins/shared/document.addhandlerasync) permettant d’ajouter un gestionnaire d’événements pour l’événement [SelectionChanged](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) sur le document.


```js
Office.context.document.addHandlerAsync("documentSelectionChanged", myHandler, function(result){} 
);

// Event handler function.
function myHandler(eventArgs){
write('Document Selection Changed');
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Le premier paramètre  _eventType_ spécifie le nom de l’événement auquel souscrire. Transmettre la chaîne `"documentSelectionChanged"` pour ce paramètre revient à transmettre le type d’événement **Office.EventType.DocumentSelectionChanged** de l’énumération [Office.EventType](https://dev.office.com/reference/add-ins/shared/eventtype-enumeration).

La fonction `myHander()` transmise dans la fonction comme deuxième paramètre _handler_ est un gestionnaire d’événements qui est exécuté lorsque la sélection change dans le document. La fonction est appelée avec un seul paramètre, _eventArgs_, qui contient une référence à un objet [DocumentSelectionChangedEventArgs](https://dev.office.com/reference/add-ins/shared/document.selectionchangedeventargs) quand l’opération asynchrone se termine. Vous pouvez utiliser la propriété [DocumentSelectionChangedEventArgs.document](https://dev.office.com/reference/add-ins/shared/document.selectionchangedeventargs.document) pour accéder au document qui a déclenché l’événement.


> [!NOTE]
> Vous pouvez ajouter plusieurs gestionnaires d’événements pour un événement donné en rappelant la méthode **addHandlerAsync** et en transmettant une fonction de gestionnaire d’événements supplémentaire au paramètre _handler_. Cela fonctionnera correctement à condition que le nom de chaque fonction de gestionnaire d’événements soit unique.


## <a name="stop-detecting-changes-in-the-selection"></a>Arrêt de la détection de modifications dans la sélection


L’exemple suivant montre comment arrêter l’écoute de l’événement [Document.SelectionChanged](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) en appelant la méthode [document.removeHandlerAsync](https://dev.office.com/reference/add-ins/shared/document.removehandlerasync).


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

Le nom de la fonction `myHandler` passé en tant que deuxième paramètre _handler_ désigne le gestionnaire d’événements qui sera supprimé de l’événement **SelectionChanged**.


> [!IMPORTANT]
> Si le paramètre facultatif _handler_ est omis lors de l’appel à la méthode **removeHandlerAsync**, tous les gestionnaires d’événements du paramètre _eventType_ spécifié seront supprimés.

