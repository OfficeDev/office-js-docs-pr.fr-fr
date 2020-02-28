---
title: Lire et écrire des données dans la sélection active d’un document ou d’une feuille de calcul
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 039631e935d2ff6fadb4eab9d99df73ac30dae4d
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325002"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>Lire et écrire des données dans la sélection active d’un document ou d’une feuille de calcul

L’objet [document](/javascript/api/office/office.document) expose des méthodes qui vous permettent de lire et d’écrire dans la sélection actuelle de l’utilisateur dans un document ou une feuille de calcul. Pour ce faire, l' `Document` objet fournit les `getSelectedDataAsync` méthodes `setSelectedDataAsync` et. Cette rubrique explique également comment lire, écrire et créer des gestionnaires d’événements pour détecter les modifications apportées à la sélection de l’utilisateur.

La `getSelectedDataAsync` méthode ne fonctionne qu’avec la sélection actuelle de l’utilisateur. Si vous devez conserver la sélection dans le document, afin que la même sélection soit disponible pour la lecture et l’écriture dans les sessions de l’exécution de votre complément, vous devez ajouter une liaison à l’aide de la méthode [bindings. addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) (ou créer une liaison avec l’une des autres méthodes « addFrom » de l’objet [bindings](/javascript/api/office/office.bindings) ). Pour plus d’informations sur la création d’une liaison à une région d’un document, puis sur la lecture et l’écriture d’une liaison, voir [lier à des régions dans un document ou une feuille de calcul](bind-to-regions-in-a-document-or-spreadsheet.md).


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

Dans cet exemple, le premier paramètre _coercionType_ est spécifié comme `Office.CoercionType.Text` (vous pouvez également spécifier ce paramètre à l’aide de la `"text"`chaîne littérale). Cela signifie que la propriété [value](/javascript/api/office/office.asyncresult#status) de l’objet [asyncResult](/javascript/api/office/office.asyncresult) qui est disponible à partir du paramètre _asyncResult_ de la fonction de rappel renverra une **chaîne** qui contient le texte sélectionné dans le document. La spécification de différents types de forçage de type entraîne des valeurs différentes. [Office. CoercionType](/javascript/api/office/office.coerciontype) est une énumération des valeurs de type de forçage de type disponibles. `Office.CoercionType.Text` prend la valeur de la chaîne « Text ».


> [!TIP]
> **Quand utiliser la matrice et le tableau coercionType pour l’accès aux données ?** Si vous souhaitez que les données de tableau sélectionnées s’étendent dynamiquement lorsque les lignes et les colonnes sont ajoutées, et que vous devez utiliser des en-têtes de tableau, vous devez utiliser le type de données table ( `"table"` en `Office.CoercionType.Table`spécifiant le paramètre _coercionType_ de la `getSelectedDataAsync` méthode comme ou). L’ajout de lignes et de colonnes dans la structure de données est pris en charge dans les données de tableau et de matrice, mais l’ajout de lignes et de colonnes est pris en charge uniquement pour les données de table. Si vous ne prévoyez pas d’ajouter des lignes et des colonnes, et que vos données ne nécessitent pas de fonctionnalité d’en-tête, vous devez utiliser le type de données Matrix `getSelectedDataAsync` (en `"matrix"` spécifiant le paramètre `Office.CoercionType.Matrix` _coercionType_ de la méthode As ou), ce qui fournit un modèle plus simple d’interaction avec les données.

La fonction anonyme qui est transmise à la fonction en tant que deuxième paramètre de _rappel_ est `getSelectedDataAsync` exécutée lorsque l’opération est terminée. La fonction est appelée avec un seul paramètre, _asyncResult_, qui contient le résultat et l’état de l’appel. En cas d’échec de l' [](/javascript/api/office/office.asyncresult#asynccontext) appel, la propriété `AsyncResult` Error de l’objet donne accès à l’objet [Error](/javascript/api/office/office.error) . Vous pouvez vérifier la valeur des propriétés [Error.Name](/javascript/api/office/office.error#name) et [Error. message](/javascript/api/office/office.error#message) afin de déterminer la cause de l’échec de l’opération set. Dans le cas contraire, le texte sélectionné dans le document est affiché.

La propriété [asyncResult. Status](/javascript/api/office/office.asyncresult#error) est utilisée dans l’instruction **If** pour vérifier si l’appel a réussi. [Office. AsyncResultStatus](/javascript/api/office/office.asyncresult#status) est une énumération des `AsyncResult.status` valeurs de propriété disponibles. `Office.AsyncResultStatus.Failed` donne la chaîne « failed » (et, à nouveau, peut également être spécifié comme cette chaîne littérale).


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

La fonction anonyme transmise à la méthode [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) en tant que paramètre _callback_ est exécutée lorsque l’appel asynchrone est terminé. Lorsque vous écrivez des données dans la sélection à l' `setSelectedDataAsync` aide de la méthode, le paramètre _asyncResult_ du rappel donne accès uniquement à l’état de l’appel et à l’objet d' [erreur](/javascript/api/office/office.error) en cas d’échec de l’appel.

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

Le premier paramètre _eventType_ spécifie le nom de l’événement auquel s’abonner. Le passage de `"documentSelectionChanged"` la chaîne pour ce paramètre équivaut à la `Office.EventType.DocumentSelectionChanged` transmission du type d’événement de l’énumération [Office. EventType](/javascript/api/office/office.eventtype) .

La fonction `myHander()` transmise dans la fonction comme deuxième paramètre _handler_ est un gestionnaire d’événements qui est exécuté lorsque la sélection change dans le document. La fonction est appelée avec un seul paramètre, _eventArgs_, qui contient une référence à un objet [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs) quand l’opération asynchrone se termine. Vous pouvez utiliser la propriété [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) pour accéder au document qui a déclenché l’événement.


> [!NOTE]
> Vous pouvez ajouter plusieurs gestionnaires d’événements pour un événement donné en appelant à `addHandlerAsync` nouveau la méthode et en transmettant une fonction de gestionnaire d’événements supplémentaire pour le paramètre _handler_ . Cela fonctionnera correctement tant que le nom de chaque fonction de gestionnaire d’événements est unique.


## <a name="stop-detecting-changes-in-the-selection"></a>Arrêt de la détection de modifications dans la sélection


L’exemple suivant montre comment arrêter l’écoute de l’événement [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) en appelant la méthode [document.removeHandlerAsync](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-).


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

Le `myHandler` nom de la fonction passé en tant que deuxième paramètre _handler_ désigne le gestionnaire d’événements qui sera supprimé de `SelectionChanged` l’événement.


> [!IMPORTANT]
> Si le paramètre facultatif _handler_ est omis lors de l' `removeHandlerAsync` appel de la méthode, tous les gestionnaires d’événements pour le paramètre _eventType_ spécifié sont supprimés.
