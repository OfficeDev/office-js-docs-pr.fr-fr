---
title: Lire et écrire des données dans la sélection active d’un document ou d’une feuille de calcul
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: d1c8fcdeec8d92fd3f77e169dc24715f7c5e9964
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944985"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a><span data-ttu-id="497f1-102">Lire et écrire des données dans la sélection active d’un document ou d’une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="497f1-102">Read and write data to the active selection in a document or spreadsheet</span></span>

<span data-ttu-id="497f1-p101">L’objet [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) expose des méthodes qui vous permettent de lire et d’écrire dans la sélection active de l’utilisateur dans un document ou une feuille de calcul. Pour cela, l’objet **Document** fournit les méthodes **getSelectedDataAsync** et **setSelectedDataAsync**. Cette rubrique explique comment lire, écrire et créer des gestionnaires d’événements pour détecter les changements intervenant dans la sélection de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="497f1-p101">The [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) object exposes methods that let you read and write to the user's current selection in a document or spreadsheet. To do that, the **Document** object provides the **getSelectedDataAsync** and **setSelectedDataAsync** methods. This topic also describes how to read, write, and create event handlers to detect changes to the user's selection.</span></span>

<span data-ttu-id="497f1-p102">La méthode **getSelectedDataAsync** ne fonctionne que sur la sélection active de l’utilisateur. Si vous devez conserver la sélection dans le document, afin que la même sélection soit disponible en lecture et en écriture dans les sessions exécutant votre complément, vous devez ajouter une liaison à l’aide de la méthode [Bindings.addFromSelectionAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromselectionasync-bindingtype--options--callback-) (ou créer une liaison à l’aide de l’une des autres méthodes « addFrom » de l’objet [Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js)). Pour plus d’informations sur la création d’une liaison vers une zone d’un document et sur la lecture et l’écriture dans une liaison, voir [Liaisons de régions dans un document ou une feuille de calcul](bind-to-regions-in-a-document-or-spreadsheet.md).</span><span class="sxs-lookup"><span data-stu-id="497f1-p102">The  **getSelectedDataAsync** method only works against the user's current selection. If you need to persist the selection in the document, so that the same selection is available to read and write across sessions of running your add-in, you must add a binding using the [Bindings.addFromSelectionAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromselectionasync-bindingtype--options--callback-) method (or create a binding with one of the other "addFrom" methods of the [Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js) object). For information about creating a binding to a region of a document, and then reading and writing to a binding, see [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span>


## <a name="read-selected-data"></a><span data-ttu-id="497f1-109">Lecture de données sélectionnées</span><span class="sxs-lookup"><span data-stu-id="497f1-109">Read selected data</span></span>


<span data-ttu-id="497f1-110">L’exemple suivant montre comment obtenir les données d’une sélection dans un document en utilisant la méthode [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="497f1-110">The following example shows how to get data from a selection in a document by using the [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) method.</span></span>


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

<span data-ttu-id="497f1-p103">Dans cet exemple, le premier paramètre _coercionType_ est spécifié comme **Office.CoercionType.Text** (vous pouvez également spécifier ce paramètre en utilisant la chaîne littérale `"text"`). Cela signifie que la propriété [value](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#status) de l’objet [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js) qui est disponible à partir du paramètre _asyncResult_ dans la fonction de rappel renverra une **string** qui contient le texte sélectionné dans le document. La spécification de différents types de forçage de type produit des valeurs différentes. [Office.CoercionType](https://docs.microsoft.com/javascript/api/office/office.coerciontype?view=office-js) est une énumération des valeurs de types de forçage de type disponibles. **Office.CoercionType.Text** prend la valeur de la chaîne « text ».</span><span class="sxs-lookup"><span data-stu-id="497f1-p103">In this example, the first  _coercionType_ parameter is specified as **Office.CoercionType.Text** (you can also specify this parameter by using the literal string `"text"`). This means that the [value](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#status) property of the [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js) object that is available from the _asyncResult_ parameter in the callback function will return a **string** that contains the selected text in the document. Specifying different coercion types will result in different values. [Office.CoercionType](https://docs.microsoft.com/javascript/api/office/office.coerciontype?view=office-js) is an enumeration of available coercion type values. **Office.CoercionType.Text** evaluates to the string "text".</span></span>


> [!TIP]
> <span data-ttu-id="497f1-p104">**Quand devez-vous utiliser la matrice ou le paramètre coercionType de tableau pour accéder aux données ?** Si les données tabulaires sélectionnées doivent croître de façon dynamique lors de l’ajout de lignes et de colonnes, et que vous devez travailler avec des en-têtes de tableaux, vous devez utiliser le type de données de tableau (en spécifiant le paramètre _coercionType_ de la méthode **getSelectedDataAsync** en tant que `"table"` ou **Office.CoercionType.Table**). L’ajout de lignes et de colonnes au sein de la structure de données est pris en charge dans les données de tableau et de matrice, mais l’ajout de lignes et de colonnes à la fin est pris en charge uniquement pour les données de tableau. Si vous ne prévoyez pas d’ajouter des lignes et des colonnes, et que vos données ne nécessitent pas la fonctionnalité d’en-tête, vous devez utiliser le type de données de matrice (en spécifiant le paramètre _coercionType_ de la méthode **getSelecteDataAsync** en tant que `"matrix"` ou **Office.CoercionType.Matrix**), qui fournit un modèle plus simple d’interaction avec les données.</span><span class="sxs-lookup"><span data-stu-id="497f1-p104">**When should you use the matrix versus table coercionType for data access?** If you need your selected tabular data to grow dynamically when rows and columns are added, and you must work with table headers, you should use the table data type (by specifying the _coercionType_ parameter of the **getSelectedDataAsync** method as `"table"` or **Office.CoercionType.Table**). Adding rows and columns within the data structure is supported in both table and matrix data, but appending rows and columns is supported only for table data. If you are you aren't planning on adding rows and columns, and your data doesn't require header functionality, then you should use the matrix data type (by specifying the  _coercionType_ parameter of **getSelecteDataAsync** method as `"matrix"` or **Office.CoercionType.Matrix**), which provides a simpler model of interacting with the data.</span></span>

<span data-ttu-id="497f1-p105">La fonction anonyme qui est transmise dans la fonction comme deuxième paramètre _callback_ est exécutée lorsque l’opération **getSelectedDataAsync** est terminée. La fonction est appelée avec un seul paramètre, _asyncResult_, qui contient le résultat et l’état de l’appel. Si l’appel échoue, la propriété [error](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#asynccontext) de l’objet **AsyncResult** donne accès à l’objet [Error](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js). Vous pouvez vérifier la valeur des propriétés [Error.name](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js#name) et [Error.message](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js#message) pour déterminer les raisons de l’échec de l’opération. Sinon, le texte sélectionné dans le document s’affiche.</span><span class="sxs-lookup"><span data-stu-id="497f1-p105">The anonymous function that is passed into the function as the second  _callback_ parameter is executed when the **getSelectedDataAsync** operation is completed. The function is called with a single parameter, _asyncResult_, which contains the result and the status of the call. If the call fails, the [error](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#asynccontext) property of the **AsyncResult** object provides access to the [Error](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js) object. You can check the value of the [Error.name](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js#name) and [Error.message](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js#message) properties to determine why the set operation failed. Otherwise, the selected text in the document is displayed.</span></span>

<span data-ttu-id="497f1-p106">La propriété [AsyncResult.status](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#error) est utilisée dans l’instruction **if** pour tester la réussite de l’appel. [Office.AsyncResultStatus](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#status) est une énumération des valeurs de propriété **AsyncResult.status** disponibles. **Office.AsyncResultStatus.Failed** prend la valeur de la chaîne « failed » (et, de nouveau, peut également être spécifié comme chaîne littérale).</span><span class="sxs-lookup"><span data-stu-id="497f1-p106">The [AsyncResult.status](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#error) property is used in the **if** statement to test whether the call succeeded. [Office.AsyncResultStatus](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js#status) is an enumeration of available **AsyncResult.status** property values. **Office.AsyncResultStatus.Failed** evaluates to the string "failed" (and, again, can also be specified as that literal string).</span></span>


## <a name="write-data-to-the-selection"></a><span data-ttu-id="497f1-128">Écriture de données dans la sélection</span><span class="sxs-lookup"><span data-stu-id="497f1-128">Write data to the selection</span></span>


<span data-ttu-id="497f1-129">L’exemple suivant montre comment définir la sélection pour afficher « Hello World! ».</span><span class="sxs-lookup"><span data-stu-id="497f1-129">The following example shows how to set the selection to show "Hello World!".</span></span>


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

<span data-ttu-id="497f1-p107">Le passage de différents types d’objets pour le paramètre  _data_ produit différents résultats. Le résultat varie en fonction de la sélection actuelle dans le document, de l’application qui héberge votre complément, et de l’éventuel passage forcé des données dans la sélection actuelle.</span><span class="sxs-lookup"><span data-stu-id="497f1-p107">Passing in different object types for the  _data_ parameter will have different results. The result depends on what is currently selected in the document, which application is hosting your add-in, and whether the data passed in can be coerced to the current selection.</span></span>

<span data-ttu-id="497f1-p108">La fonction anonyme transmise dans la méthode [setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#setselecteddataasync-data--options--callback-) comme paramètre _callback_ est exécutée quand l’appel anonyme est terminé. Lorsque vous écrivez des données dans la sélection à l’aide de la méthode **setSelectedDataAsync**, le paramètre _asyncResult_ du rappel donne uniquement accès à l’état de l’appel et à l’objet [Error](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js) si l’appel échoue.</span><span class="sxs-lookup"><span data-stu-id="497f1-p108">The anonymous function passed into the [setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#setselecteddataasync-data--options--callback-) method as the _callback_ parameter is executed when the asynchronous call is completed. When you write data to the selection by using the **setSelectedDataAsync** method, the _asyncResult_ parameter of the callback provides access only to the status of the call, and to the [Error](https://docs.microsoft.com/javascript/api/office/office.error?view=office-js) object if the call fails.</span></span>

> [!NOTE]
> <span data-ttu-id="497f1-134">Depuis la publication d’Excel 2013 SP1 et de la version correspondante d’Excel Online, vous pouvez désormais [définir la mise en forme lors de l’écriture d’un tableau sur la sélection active](../excel/excel-add-ins-tables.md).</span><span class="sxs-lookup"><span data-stu-id="497f1-134">Starting with the release of the Excel 2013 SP1 and the corresponding build of Excel Online, you can now [set formatting when writing a table to the current selection](../excel/excel-add-ins-tables.md).</span></span>


## <a name="detect-changes-in-the-selection"></a><span data-ttu-id="497f1-135">Détection de modifications dans la sélection</span><span class="sxs-lookup"><span data-stu-id="497f1-135">Detect changes in the selection</span></span>


<span data-ttu-id="497f1-136">L’exemple suivant montre comment détecter des modifications dans la sélection à l’aide de la méthode [Document.addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) permettant d’ajouter un gestionnaire d’événements pour l’événement [SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) sur le document.</span><span class="sxs-lookup"><span data-stu-id="497f1-136">The following example shows how to detect changes in the selection by using the [Document.addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) method to add an event handler for the [SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) event on the document.</span></span>


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

<span data-ttu-id="497f1-p109">Le premier paramètre  _eventType_ spécifie le nom de l’événement auquel souscrire. Transmettre la chaîne `"documentSelectionChanged"` pour ce paramètre revient à transmettre le type d’événement **Office.EventType.DocumentSelectionChanged** de l’énumération [Office.EventType](https://docs.microsoft.com/javascript/api/office/office.eventtype?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="497f1-p109">The first  _eventType_ parameter specifies the name of the event to subscribe to. Passing the string `"documentSelectionChanged"` for this parameter is equivalent to passing the **Office.EventType.DocumentSelectionChanged** event type of the [Office.EventType](https://docs.microsoft.com/javascript/api/office/office.eventtype?view=office-js) enumeration.</span></span>

<span data-ttu-id="497f1-p110">La fonction `myHander()` transmise dans la fonction comme deuxième paramètre _handler_ est un gestionnaire d’événements qui est exécuté lorsque la sélection change dans le document. La fonction est appelée avec un seul paramètre, _eventArgs_, qui contient une référence à un objet [DocumentSelectionChangedEventArgs](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) quand l’opération asynchrone se termine. Vous pouvez utiliser la propriété [DocumentSelectionChangedEventArgs.document](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js#document) pour accéder au document qui a déclenché l’événement.</span><span class="sxs-lookup"><span data-stu-id="497f1-p110">The  `myHander()` function that is passed into the function as the second _handler_ parameter is an event handler that is executed when the selection is changed on the document. The function is called with a single parameter, _eventArgs_, which will contain a reference to a [DocumentSelectionChangedEventArgs](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) object when the asynchronous operation completes. You can use the [DocumentSelectionChangedEventArgs.document](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js#document) property to access the document that raised the event.</span></span>


> [!NOTE]
> <span data-ttu-id="497f1-p111">Vous pouvez ajouter plusieurs gestionnaires d’événements pour un événement donné en rappelant la méthode **addHandlerAsync** et en transmettant une fonction de gestionnaire d’événements supplémentaire au paramètre _handler_. Cela fonctionnera correctement à condition que le nom de chaque fonction de gestionnaire d’événements soit unique.</span><span class="sxs-lookup"><span data-stu-id="497f1-p111">You can add multiple event handlers for a given event by calling the  **addHandlerAsync** method again and passing in an additional event handler function for the _handler_ parameter. This will work correctly as long as the name of each event handler function is unique.</span></span>


## <a name="stop-detecting-changes-in-the-selection"></a><span data-ttu-id="497f1-144">Arrêt de la détection de modifications dans la sélection</span><span class="sxs-lookup"><span data-stu-id="497f1-144">Stop detecting changes in the selection</span></span>


<span data-ttu-id="497f1-145">L’exemple suivant montre comment arrêter l’écoute de l’événement [Document.SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) en appelant la méthode [document.removeHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#removehandlerasync-eventtype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="497f1-145">The following example shows how to stop listening to the [Document.SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) event by calling the [document.removeHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#removehandlerasync-eventtype--options--callback-) method.</span></span>


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

<span data-ttu-id="497f1-146">Le nom de la fonction `myHandler` passé en tant que deuxième paramètre _handler_ désigne le gestionnaire d’événements qui sera supprimé de l’événement **SelectionChanged**.</span><span class="sxs-lookup"><span data-stu-id="497f1-146">The  `myHandler` function name that is passed as the second _handler_ parameter specifies the event handler that will be removed from the **SelectionChanged** event.</span></span>


> [!IMPORTANT]
> <span data-ttu-id="497f1-147">Si le paramètre facultatif _handler_ est omis lors de l’appel à la méthode **removeHandlerAsync**, tous les gestionnaires d’événements du paramètre _eventType_ spécifié seront supprimés.</span><span class="sxs-lookup"><span data-stu-id="497f1-147">If the optional  _handler_ parameter is omitted when the **removeHandlerAsync** method is called, all event handlers for the specified _eventType_ will be removed.</span></span>

