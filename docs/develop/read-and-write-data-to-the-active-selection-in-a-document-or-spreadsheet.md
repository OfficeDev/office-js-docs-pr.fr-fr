---
title: Lire et écrire des données dans la sélection active d’un document ou d’une feuille de calcul
description: Découvrez comment lire et écrire des données dans la sélection active dans un document Word ou Excel feuille de calcul.
ms.date: 06/20/2019
ms.localizationpriority: medium
ms.openlocfilehash: c8a199d5c6491f91a13c61a9b87ab6f302be9105
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149993"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>Lecture et écriture de données dans la sélection active d’un document ou d’une feuille de calcul

L’objet [Document](/javascript/api/office/office.document) expose des méthodes qui vous permettent de lire et d’écrire dans la sélection active de l’utilisateur dans un document ou une feuille de calcul. Pour ce faire, `Document` l’objet fournit les `getSelectedDataAsync` méthodes et les `setSelectedDataAsync` méthodes. Cette rubrique explique comment lire, écrire et créer des gestionnaires d’événements pour détecter les changements intervenant dans la sélection de l’utilisateur.

La `getSelectedDataAsync` méthode ne fonctionne que par rapport à la sélection actuelle de l’utilisateur. Si vous devez conserver la sélection dans le document, afin que la même sélection soit disponible en lecture et en écriture dans les sessions exécutant votre complément, vous devez ajouter une liaison à l’aide de la méthode [Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#addFromSelectionAsync_bindingType__options__callback_) (ou créer une liaison à l’aide de l’une des autres méthodes « addFrom » de l’objet [Bindings](/javascript/api/office/office.bindings)). Pour plus d’informations sur la création d’une liaison vers une zone d’un document et sur la lecture et l’écriture dans une liaison, voir [Liaisons de régions dans un document ou une feuille de calcul](bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="read-selected-data"></a>Lecture de données sélectionnées


L’exemple suivant montre comment obtenir les données d’une sélection dans un document en utilisant la méthode [getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_).


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

Dans cet exemple, le premier  _paramètre coercionType_ est spécifié comme (vous pouvez également spécifier ce paramètre à l’aide de `Office.CoercionType.Text` la chaîne `"text"` littérale). Cela signifie que la propriété [value](/javascript/api/office/office.asyncresult#status) de l’objet [AsyncResult](/javascript/api/office/office.asyncresult) qui est disponible à partir du paramètre _asyncResult_ dans la fonction de rappel renverra une **string** qui contient le texte sélectionné dans le document. La spécification de différents types de forçage de type produit des valeurs différentes. [Office.CoercionType](/javascript/api/office/office.coerciontype) est une énumération des valeurs de types de forçage de type disponibles. `Office.CoercionType.Text` est évaluée à la chaîne « text ».


> [!TIP]
> **Quand devez-vous utiliser la matrice ou le paramètre coercionType de tableau pour accéder aux données ?** Si vous avez besoin que vos données tabulaires sélectionnées s’développent dynamiquement lorsque des lignes et des colonnes sont ajoutées et que vous devez utiliser des en-têtes de tableau, vous devez utiliser le type de données de tableau (en spécifiant le paramètre _coercionType_ de la méthode en tant que ou `getSelectedDataAsync` `"table"` `Office.CoercionType.Table` ). L’ajout de lignes et de colonnes au sein de la structure de données est pris en charge dans les données de tableau et de matrice, mais l’ajout de lignes et de colonnes à la fin est pris en charge uniquement pour les données de tableau. Si vous n’envisagez pas d’ajouter des lignes et des colonnes et que vos données ne nécessitent pas de fonctionnalité d’en-tête, vous devez utiliser le type de données de matrice (en spécifiant le paramètre  _coercionType_ de la méthode en tant que ou ), ce qui fournit un modèle plus simple d’interaction avec les `getSelectedDataAsync` `"matrix"` `Office.CoercionType.Matrix` données.

La fonction anonyme qui est passée dans la fonction en tant que  _deuxième_ paramètre de rappel est exécutée lorsque l’opération `getSelectedDataAsync` est terminée. La fonction est appelée avec un seul paramètre, _asyncResult_, qui contient le résultat et l’état de l’appel. Si l’appel échoue, la [propriété d’erreur](/javascript/api/office/office.asyncresult#error) de l’objet `AsyncResult` donne accès à [l’objet Error.](/javascript/api/office/office.error) Vous pouvez vérifier la valeur des propriétés [Error.name](/javascript/api/office/office.error#name) et [Error.message](/javascript/api/office/office.error#message) pour déterminer les raisons de l’échec de l’opération. Sinon, le texte sélectionné dans le document s’affiche.

La propriété [AsyncResult.status](/javascript/api/office/office.asyncresult#error) est utilisée dans l’instruction **if** pour tester la réussite de l’appel. [Office. AsyncResultStatus est](/javascript/api/office/office.asyncresult#status) une éumération des valeurs de `AsyncResult.status` propriété disponibles. `Office.AsyncResultStatus.Failed` est évaluée à la chaîne « failed » (et, là encore, peut également être spécifiée en tant que chaîne littérale).


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

La fonction anonyme transmise dans la méthode [setSelectedDataAsync](/javascript/api/office/office.document#setSelectedDataAsync_data__options__callback_) comme paramètre _callback_ est exécutée quand l’appel anonyme est terminé. Lorsque vous écrivez des données dans la sélection à l’aide de la méthode, le paramètre `setSelectedDataAsync` _asyncResult_ du rappel fournit un accès uniquement à l’état de l’appel et à l’objet [Error](/javascript/api/office/office.error) en cas d’échec de l’appel.

> [!NOTE]
> Depuis la publication d’Excel 2013 SP1 et de la version correspondante d’Excel sur le web, vous pouvez désormais [définir la mise en forme lors de l’écriture d’un tableau sur la sélection active](../excel/excel-add-ins-tables.md).


## <a name="detect-changes-in-the-selection"></a>Détection de modifications dans la sélection


L’exemple suivant montre comment détecter des modifications dans la sélection à l’aide de la méthode [Document.addHandlerAsync](/javascript/api/office/office.document#addHandlerAsync_eventType__handler__options__callback_) permettant d’ajouter un gestionnaire d’événements pour l’événement [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) sur le document.


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

Le premier paramètre  _eventType_ spécifie le nom de l’événement auquel souscrire. Transmettre la chaîne pour ce paramètre équivaut à transmettre le type d’événement `"documentSelectionChanged"` `Office.EventType.DocumentSelectionChanged` du [Office. EventType,](/javascript/api/office/office.eventtype) éumération.

La fonction `myHander()` transmise dans la fonction comme deuxième paramètre _handler_ est un gestionnaire d’événements qui est exécuté lorsque la sélection change dans le document. La fonction est appelée avec un seul paramètre, _eventArgs_, qui contient une référence à un objet [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs) quand l’opération asynchrone se termine. Vous pouvez utiliser la propriété [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) pour accéder au document qui a déclenché l’événement.


> [!NOTE]
> Vous pouvez ajouter plusieurs handlers d’événements pour un événement donné en appelant à nouveau la méthode et en passant une fonction de handler d’événement supplémentaire pour le `addHandlerAsync` _paramètre de_ handler. Cela fonctionnera correctement à condition que le nom de chaque fonction de gestionnaire d’événements soit unique.


## <a name="stop-detecting-changes-in-the-selection"></a>Arrêt de la détection de modifications dans la sélection


L’exemple suivant montre comment arrêter l’écoute de l’événement [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) en appelant la méthode [document.removeHandlerAsync](/javascript/api/office/office.document#removeHandlerAsync_eventType__options__callback_).


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

Le nom de la fonction transmis en tant que deuxième paramètre de handler spécifie le handler d’événements qui `myHandler` sera supprimé de  l’événement. `SelectionChanged`


> [!IMPORTANT]
> Si le paramètre _de handler_ facultatif est omis lorsque la méthode est appelée, tous les handlers d’événements pour le type d’événement spécifié `removeHandlerAsync` sont supprimés. 
