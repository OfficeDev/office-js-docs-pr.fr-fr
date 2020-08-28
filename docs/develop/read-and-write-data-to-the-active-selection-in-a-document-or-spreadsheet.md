---
title: Lire et écrire des données dans la sélection active d’un document ou d’une feuille de calcul
description: Découvrez comment lire et écrire des données dans la sélection active dans un document Word ou une feuille de calcul Excel.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 9eaf0aac406731a9c0033e69bd8946464a4d1a4f
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292742"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>Lecture et écriture de données dans la sélection active d’un document ou d’une feuille de calcul

L’objet [Document](/javascript/api/office/office.document) expose des méthodes qui vous permettent de lire et d’écrire dans la sélection active de l’utilisateur dans un document ou une feuille de calcul. Pour ce faire, l' `Document` objet fournit les `getSelectedDataAsync` `setSelectedDataAsync` méthodes et. Cette rubrique explique comment lire, écrire et créer des gestionnaires d’événements pour détecter les changements intervenant dans la sélection de l’utilisateur.

La `getSelectedDataAsync` méthode ne fonctionne qu’avec la sélection actuelle de l’utilisateur. Si vous devez conserver la sélection dans le document, afin que la même sélection soit disponible en lecture et en écriture dans les sessions exécutant votre complément, vous devez ajouter une liaison à l’aide de la méthode [Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) (ou créer une liaison à l’aide de l’une des autres méthodes « addFrom » de l’objet [Bindings](/javascript/api/office/office.bindings)). Pour plus d’informations sur la création d’une liaison vers une zone d’un document et sur la lecture et l’écriture dans une liaison, voir [Liaisons de régions dans un document ou une feuille de calcul](bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="read-selected-data"></a>Lecture de données sélectionnées


L’exemple suivant montre comment obtenir les données d’une sélection dans un document en utilisant la méthode [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-).


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

Dans cet exemple, le premier paramètre  _coercionType_ est spécifié comme `Office.CoercionType.Text` (vous pouvez également spécifier ce paramètre à l’aide de la chaîne littérale `"text"` ). Cela signifie que la propriété [value](/javascript/api/office/office.asyncresult#status) de l’objet [AsyncResult](/javascript/api/office/office.asyncresult) qui est disponible à partir du paramètre _asyncResult_ dans la fonction de rappel renverra une **string** qui contient le texte sélectionné dans le document. La spécification de différents types de forçage de type produit des valeurs différentes. [Office.CoercionType](/javascript/api/office/office.coerciontype) est une énumération des valeurs de types de forçage de type disponibles. `Office.CoercionType.Text` prend la valeur de la chaîne « Text ».


> [!TIP]
> **Quand devez-vous utiliser la matrice ou le paramètre coercionType de tableau pour accéder aux données ?** Si vous souhaitez que les données de tableau sélectionnées s’étendent dynamiquement lorsque les lignes et les colonnes sont ajoutées, et que vous devez utiliser des en-têtes de tableau, vous devez utiliser le type de données table (en spécifiant le paramètre _coercionType_ de la `getSelectedDataAsync` méthode comme `"table"` ou `Office.CoercionType.Table` ). L’ajout de lignes et de colonnes au sein de la structure de données est pris en charge dans les données de tableau et de matrice, mais l’ajout de lignes et de colonnes à la fin est pris en charge uniquement pour les données de tableau. Si vous ne prévoyez pas d’ajouter des lignes et des colonnes, et que vos données ne nécessitent pas de fonctionnalité d’en-tête, vous devez utiliser le type de données Matrix (en spécifiant le paramètre  _coercionType_ de `getSelectedDataAsync` la méthode As `"matrix"` ou `Office.CoercionType.Matrix` ), ce qui fournit un modèle plus simple d’interaction avec les données.

La fonction anonyme qui est transmise à la fonction en tant que deuxième paramètre de  _rappel_ est exécutée lorsque l' `getSelectedDataAsync` opération est terminée. La fonction est appelée avec un seul paramètre, _asyncResult_, qui contient le résultat et l’état de l’appel. En cas d’échec de l’appel, la propriété [Error](/javascript/api/office/office.asyncresult#asynccontext) de l' `AsyncResult` objet donne accès à l’objet [Error](/javascript/api/office/office.error) . Vous pouvez vérifier la valeur des propriétés [Error.name](/javascript/api/office/office.error#name) et [Error.message](/javascript/api/office/office.error#message) pour déterminer les raisons de l’échec de l’opération. Sinon, le texte sélectionné dans le document s’affiche.

La propriété [AsyncResult.status](/javascript/api/office/office.asyncresult#error) est utilisée dans l’instruction **if** pour tester la réussite de l’appel. [Office. AsyncResultStatus](/javascript/api/office/office.asyncresult#status) est une énumération des `AsyncResult.status` valeurs de propriété disponibles. `Office.AsyncResultStatus.Failed` donne la chaîne « failed » (et, à nouveau, peut également être spécifié comme cette chaîne littérale).


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

Le passage de différents types d’objets pour le paramètre  _data_ produit différents résultats. Le résultat dépend de ce qui est actuellement sélectionné dans le document, de l’application cliente Office qui héberge votre complément et du fait que les données transmises peuvent être forcées ou non à la sélection actuelle.

La fonction anonyme transmise dans la méthode [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) comme paramètre _callback_ est exécutée quand l’appel anonyme est terminé. Lorsque vous écrivez des données dans la sélection à l’aide de la `setSelectedDataAsync` méthode, le paramètre _asyncResult_ du rappel donne accès uniquement à l’état de l’appel et à l’objet d' [erreur](/javascript/api/office/office.error) en cas d’échec de l’appel.

> [!NOTE]
> Depuis la publication d’Excel 2013 SP1 et de la version correspondante d’Excel sur le web, vous pouvez désormais [définir la mise en forme lors de l’écriture d’un tableau sur la sélection active](../excel/excel-add-ins-tables.md).


## <a name="detect-changes-in-the-selection"></a>Détection de modifications dans la sélection


L’exemple suivant montre comment détecter des modifications dans la sélection à l’aide de la méthode [Document.addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) permettant d’ajouter un gestionnaire d’événements pour l’événement [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) sur le document.


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

Le premier paramètre  _eventType_ spécifie le nom de l’événement auquel souscrire. Le passage de la chaîne `"documentSelectionChanged"` pour ce paramètre équivaut à la transmission du `Office.EventType.DocumentSelectionChanged` type d’événement de l’énumération [Office. EventType](/javascript/api/office/office.eventtype) .

La fonction `myHander()` transmise dans la fonction comme deuxième paramètre _handler_ est un gestionnaire d’événements qui est exécuté lorsque la sélection change dans le document. La fonction est appelée avec un seul paramètre, _eventArgs_, qui contient une référence à un objet [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs) quand l’opération asynchrone se termine. Vous pouvez utiliser la propriété [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) pour accéder au document qui a déclenché l’événement.


> [!NOTE]
> Vous pouvez ajouter plusieurs gestionnaires d’événements pour un événement donné en appelant `addHandlerAsync` à nouveau la méthode et en transmettant une fonction de gestionnaire d’événements supplémentaire pour le paramètre _handler_ . Cela fonctionnera correctement à condition que le nom de chaque fonction de gestionnaire d’événements soit unique.


## <a name="stop-detecting-changes-in-the-selection"></a>Arrêt de la détection de modifications dans la sélection


L’exemple suivant montre comment arrêter l’écoute de l’événement [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) en appelant la méthode [document.removeHandlerAsync](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-).


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

Le  `myHandler` nom de la fonction passé en tant que deuxième paramètre _handler_ désigne le gestionnaire d’événements qui sera supprimé de l' `SelectionChanged` événement.


> [!IMPORTANT]
> Si le paramètre facultatif  _handler_ est omis lors de l' `removeHandlerAsync` appel de la méthode, tous les gestionnaires d’événements pour le paramètre _eventType_ spécifié sont supprimés.
