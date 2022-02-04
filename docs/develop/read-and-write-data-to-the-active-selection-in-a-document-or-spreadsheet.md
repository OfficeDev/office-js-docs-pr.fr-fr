---
title: Lire et écrire des données dans la sélection active d’un document ou d’une feuille de calcul
description: Découvrez comment lire et écrire des données dans la sélection active dans un document Word ou Excel feuille de calcul.
ms.date: 01/31/2022
ms.localizationpriority: medium
---


# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>Lecture et écriture de données dans la sélection active d’un document ou d’une feuille de calcul

L’objet [Document](/javascript/api/office/office.document) expose des méthodes qui vous permettent de lire et d’écrire dans la sélection active de l’utilisateur dans un document ou une feuille de calcul. Pour ce faire, l’objet `Document` fournit les méthodes `getSelectedDataAsync` `setSelectedDataAsync` et les méthodes. Cette rubrique explique comment lire, écrire et créer des gestionnaires d’événements pour détecter les changements intervenant dans la sélection de l’utilisateur.

La `getSelectedDataAsync` méthode ne fonctionne que par rapport à la sélection actuelle de l’utilisateur. Si vous devez conserver la sélection dans le document, afin que la même sélection soit disponible en lecture et en écriture dans les sessions exécutant votre complément, vous devez ajouter une liaison à l’aide de la méthode [Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1)) (ou créer une liaison à l’aide de l’une des autres méthodes « addFrom » de l’objet [Bindings](/javascript/api/office/office.bindings)). Pour plus d’informations sur la création d’une liaison vers une zone d’un document et sur la lecture et l’écriture dans une liaison, voir [Liaisons de régions dans un document ou une feuille de calcul](bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="read-selected-data"></a>Lecture de données sélectionnées


L’exemple suivant montre comment obtenir les données d’une sélection dans un document en utilisant la méthode [getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)).


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

Dans cet exemple, le premier  _paramètre coercionType_ `Office.CoercionType.Text` est spécifié comme (vous pouvez également spécifier ce paramètre à l’aide de la chaîne littérale `"text"`). Cela signifie que la propriété [value](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) de l’objet [AsyncResult](/javascript/api/office/office.asyncresult) qui est disponible à partir du paramètre _asyncResult_ dans la fonction de rappel renverra une **string** qui contient le texte sélectionné dans le document. La spécification de différents types de forçage de type produit des valeurs différentes. [Office.CoercionType](/javascript/api/office/office.coerciontype) est une énumération des valeurs de types de forçage de type disponibles. `Office.CoercionType.Text` est évaluée à la chaîne « text ».


> [!TIP]
> **Quand devez-vous utiliser la matrice ou le paramètre coercionType de tableau pour accéder aux données ?** Si vous avez besoin que vos données tabulaires sélectionnées s’développent dynamiquement lorsque des lignes et des colonnes sont ajoutées et que vous devez utiliser des en-têtes de tableau, vous devez utiliser le type de données de tableau (en spécifiant le paramètre _coercionType_ `getSelectedDataAsync` `"table"` `Office.CoercionType.Table`de la méthode en tant que ou ). L’ajout de lignes et de colonnes au sein de la structure de données est pris en charge dans les données de tableau et de matrice, mais l’ajout de lignes et de colonnes à la fin est pris en charge uniquement pour les données de tableau. Si vous n’envisagez pas d’ajouter des lignes et des colonnes et que vos données ne nécessitent pas de fonctionnalité d’en-tête, vous devez utiliser le type de données de matrice (en spécifiant le paramètre  _coercionType_ `getSelectedDataAsync` de la méthode en `"matrix"` `Office.CoercionType.Matrix`tant que ou ), ce qui fournit un modèle plus simple d’interaction avec les données.

La fonction anonyme qui est passée dans la fonction en tant  _que deuxième paramètre_ de rappel est exécutée lorsque `getSelectedDataAsync` l’opération est terminée. La fonction est appelée avec un seul paramètre, _asyncResult_, qui contient le résultat et l’état de l’appel. Si l’appel échoue, la [propriété d’erreur](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) de l’objet `AsyncResult` donne accès à [l’objet Error](/javascript/api/office/office.error) . Vous pouvez vérifier la valeur des propriétés [Error.name](/javascript/api/office/office.error#office-office-error-name-member) et [Error.message](/javascript/api/office/office.error#office-office-error-message-member) pour déterminer les raisons de l’échec de l’opération. Sinon, le texte sélectionné dans le document s’affiche.

La propriété [AsyncResult.status](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) est utilisée dans l’instruction **if** pour tester la réussite de l’appel. [Office. AsyncResultStatus est](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) une éumération des valeurs de `AsyncResult.status` propriété disponibles. `Office.AsyncResultStatus.Failed` est évaluée à la chaîne « failed » (et, là encore, peut également être spécifiée en tant que chaîne littérale).


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

Le passage de différents types d’objets pour le paramètre  _data_ produit différents résultats. Le résultat dépend de ce qui est actuellement sélectionné dans le document, de l’application cliente Office héberge votre add-in et de la sélection actuelle des données transmises.

La fonction anonyme transmise dans la méthode [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) comme paramètre _callback_ est exécutée quand l’appel anonyme est terminé. `setSelectedDataAsync` Lorsque vous écrivez des données dans la sélection à l’aide de la méthode, le paramètre _asyncResult_ du rappel fournit un accès uniquement à l’état de l’appel et à l’objet [Error](/javascript/api/office/office.error) en cas d’échec de l’appel.

> [!NOTE]
> Depuis la publication d’Excel 2013 SP1 et de la version correspondante d’Excel sur le web, vous pouvez désormais [définir la mise en forme lors de l’écriture d’un tableau sur la sélection active](../excel/excel-add-ins-tables.md).


## <a name="detect-changes-in-the-selection"></a>Détection de modifications dans la sélection


L’exemple suivant montre comment détecter des modifications dans la sélection à l’aide de la méthode [Document.addHandlerAsync](/javascript/api/office/office.document#office-office-document-addhandlerasync-member(1)) permettant d’ajouter un gestionnaire d’événements pour l’événement [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) sur le document.


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

Le premier paramètre  _eventType_ spécifie le nom de l’événement auquel souscrire. Transmettre la chaîne `"documentSelectionChanged"` pour ce paramètre équivaut à transmettre le `Office.EventType.DocumentSelectionChanged` type d’événement du [Office. EventType](/javascript/api/office/office.eventtype), éumération.

La fonction `myHandler()` transmise dans la fonction comme deuxième paramètre _handler_ est un gestionnaire d’événements qui est exécuté lorsque la sélection change dans le document. La fonction est appelée avec un seul paramètre, _eventArgs_, qui contient une référence à un objet [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs) quand l’opération asynchrone se termine. Vous pouvez utiliser la propriété [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#office-office-documentselectionchangedeventargs-document-member) pour accéder au document qui a déclenché l’événement.


> [!NOTE]
> Vous pouvez ajouter plusieurs handlers d’événements `addHandlerAsync` pour un événement donné en appelant à nouveau la méthode et en passant une fonction de handler d’événements supplémentaire pour le _paramètre de_ handler. Cela fonctionnera correctement à condition que le nom de chaque fonction de gestionnaire d’événements soit unique.


## <a name="stop-detecting-changes-in-the-selection"></a>Arrêt de la détection de modifications dans la sélection


L’exemple suivant montre comment arrêter l’écoute de l’événement [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) en appelant la méthode [document.removeHandlerAsync](/javascript/api/office/office.document#office-office-document-removehandlerasync-member(1)).


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

Le  `myHandler` nom de la fonction transmis en tant que deuxième paramètre _de handler_ spécifie le handler d’événements qui sera supprimé de l’événement `SelectionChanged` .


> [!IMPORTANT]
> Si le paramètre  _de handler_ facultatif `removeHandlerAsync` est omis lorsque la méthode est appelée, tous les handlers d’événements pour le _type_ d’événement spécifié sont supprimés.
