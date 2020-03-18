---
title: Lire et écrire des données dans la sélection active d’un document ou d’une feuille de calcul
description: Découvrez comment lire et écrire des données dans la sélection active dans un document Word ou une feuille de calcul Excel.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 83f3de5c522436ac06a0238781ee71de676297a1
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718880"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a><span data-ttu-id="c375b-103">Lecture et écriture de données dans la sélection active d’un document ou d’une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="c375b-103">Read and write data to the active selection in a document or spreadsheet</span></span>

<span data-ttu-id="c375b-104">L’objet [Document](/javascript/api/office/office.document) expose des méthodes qui vous permettent de lire et d’écrire dans la sélection active de l’utilisateur dans un document ou une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="c375b-104">The [Document](/javascript/api/office/office.document) object exposes methods that let you read and write to the user's current selection in a document or spreadsheet.</span></span> <span data-ttu-id="c375b-105">Pour ce faire, l' `Document` objet fournit les `getSelectedDataAsync` méthodes `setSelectedDataAsync` et.</span><span class="sxs-lookup"><span data-stu-id="c375b-105">To do that, the `Document` object provides the `getSelectedDataAsync` and `setSelectedDataAsync` methods.</span></span> <span data-ttu-id="c375b-106">Cette rubrique explique comment lire, écrire et créer des gestionnaires d’événements pour détecter les changements intervenant dans la sélection de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c375b-106">This topic also describes how to read, write, and create event handlers to detect changes to the user's selection.</span></span>

<span data-ttu-id="c375b-107">La `getSelectedDataAsync` méthode ne fonctionne qu’avec la sélection actuelle de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c375b-107">The `getSelectedDataAsync` method only works against the user's current selection.</span></span> <span data-ttu-id="c375b-108">Si vous devez conserver la sélection dans le document, afin que la même sélection soit disponible en lecture et en écriture dans les sessions exécutant votre complément, vous devez ajouter une liaison à l’aide de la méthode [Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) (ou créer une liaison à l’aide de l’une des autres méthodes « addFrom » de l’objet [Bindings](/javascript/api/office/office.bindings)).</span><span class="sxs-lookup"><span data-stu-id="c375b-108">If you need to persist the selection in the document, so that the same selection is available to read and write across sessions of running your add-in, you must add a binding using the [Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) method (or create a binding with one of the other "addFrom" methods of the [Bindings](/javascript/api/office/office.bindings) object).</span></span> <span data-ttu-id="c375b-109">Pour plus d’informations sur la création d’une liaison vers une zone d’un document et sur la lecture et l’écriture dans une liaison, voir [Liaisons de régions dans un document ou une feuille de calcul](bind-to-regions-in-a-document-or-spreadsheet.md).</span><span class="sxs-lookup"><span data-stu-id="c375b-109">For information about creating a binding to a region of a document, and then reading and writing to a binding, see [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span>


## <a name="read-selected-data"></a><span data-ttu-id="c375b-110">Lecture de données sélectionnées</span><span class="sxs-lookup"><span data-stu-id="c375b-110">Read selected data</span></span>


<span data-ttu-id="c375b-111">L’exemple suivant montre comment obtenir les données d’une sélection dans un document en utilisant la méthode [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="c375b-111">The following example shows how to get data from a selection in a document by using the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) method.</span></span>


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

<span data-ttu-id="c375b-112">Dans cet exemple, le premier paramètre _coercionType_ est spécifié comme `Office.CoercionType.Text` (vous pouvez également spécifier ce paramètre à l’aide de la `"text"`chaîne littérale).</span><span class="sxs-lookup"><span data-stu-id="c375b-112">In this example, the first  _coercionType_ parameter is specified as `Office.CoercionType.Text` (you can also specify this parameter by using the literal string `"text"`).</span></span> <span data-ttu-id="c375b-113">Cela signifie que la propriété [value](/javascript/api/office/office.asyncresult#status) de l’objet [AsyncResult](/javascript/api/office/office.asyncresult) qui est disponible à partir du paramètre _asyncResult_ dans la fonction de rappel renverra une **string** qui contient le texte sélectionné dans le document.</span><span class="sxs-lookup"><span data-stu-id="c375b-113">This means that the [value](/javascript/api/office/office.asyncresult#status) property of the [AsyncResult](/javascript/api/office/office.asyncresult) object that is available from the _asyncResult_ parameter in the callback function will return a **string** that contains the selected text in the document.</span></span> <span data-ttu-id="c375b-114">La spécification de différents types de forçage de type produit des valeurs différentes.</span><span class="sxs-lookup"><span data-stu-id="c375b-114">Specifying different coercion types will result in different values.</span></span> <span data-ttu-id="c375b-115">[Office.CoercionType](/javascript/api/office/office.coerciontype) est une énumération des valeurs de types de forçage de type disponibles.</span><span class="sxs-lookup"><span data-stu-id="c375b-115">[Office.CoercionType](/javascript/api/office/office.coerciontype) is an enumeration of available coercion type values.</span></span> <span data-ttu-id="c375b-116">`Office.CoercionType.Text`prend la valeur de la chaîne « Text ».</span><span class="sxs-lookup"><span data-stu-id="c375b-116">`Office.CoercionType.Text` evaluates to the string "text".</span></span>


> [!TIP]
> <span data-ttu-id="c375b-117">**Quand devez-vous utiliser la matrice ou le paramètre coercionType de tableau pour accéder aux données ?**</span><span class="sxs-lookup"><span data-stu-id="c375b-117">**When should you use the matrix versus table coercionType for data access?**</span></span> <span data-ttu-id="c375b-118">Si vous souhaitez que les données de tableau sélectionnées s’étendent dynamiquement lorsque les lignes et les colonnes sont ajoutées, et que vous devez utiliser des en-têtes de tableau, vous devez utiliser le type de données table ( `"table"` en `Office.CoercionType.Table`spécifiant le paramètre _coercionType_ de la `getSelectedDataAsync` méthode comme ou).</span><span class="sxs-lookup"><span data-stu-id="c375b-118">If you need your selected tabular data to grow dynamically when rows and columns are added, and you must work with table headers, you should use the table data type (by specifying the _coercionType_ parameter of the `getSelectedDataAsync` method as `"table"` or `Office.CoercionType.Table`).</span></span> <span data-ttu-id="c375b-119">L’ajout de lignes et de colonnes au sein de la structure de données est pris en charge dans les données de tableau et de matrice, mais l’ajout de lignes et de colonnes à la fin est pris en charge uniquement pour les données de tableau.</span><span class="sxs-lookup"><span data-stu-id="c375b-119">Adding rows and columns within the data structure is supported in both table and matrix data, but appending rows and columns is supported only for table data.</span></span> <span data-ttu-id="c375b-120">Si vous ne prévoyez pas d’ajouter des lignes et des colonnes, et que vos données ne nécessitent pas de fonctionnalité d’en-tête, vous devez utiliser le type de données Matrix `getSelectedDataAsync` (en `"matrix"` spécifiant le paramètre `Office.CoercionType.Matrix` _coercionType_ de la méthode As ou), ce qui fournit un modèle plus simple d’interaction avec les données.</span><span class="sxs-lookup"><span data-stu-id="c375b-120">If you are you aren't planning on adding rows and columns, and your data doesn't require header functionality, then you should use the matrix data type (by specifying the  _coercionType_ parameter of `getSelectedDataAsync` method as `"matrix"` or `Office.CoercionType.Matrix`), which provides a simpler model of interacting with the data.</span></span>

<span data-ttu-id="c375b-121">La fonction anonyme qui est transmise à la fonction en tant que deuxième paramètre de _rappel_ est `getSelectedDataAsync` exécutée lorsque l’opération est terminée.</span><span class="sxs-lookup"><span data-stu-id="c375b-121">The anonymous function that is passed into the function as the second  _callback_ parameter is executed when the `getSelectedDataAsync` operation is completed.</span></span> <span data-ttu-id="c375b-122">La fonction est appelée avec un seul paramètre, _asyncResult_, qui contient le résultat et l’état de l’appel.</span><span class="sxs-lookup"><span data-stu-id="c375b-122">The function is called with a single parameter, _asyncResult_, which contains the result and the status of the call.</span></span> <span data-ttu-id="c375b-123">En cas d’échec de l' [error](/javascript/api/office/office.asyncresult#asynccontext) appel, la propriété `AsyncResult` Error de l’objet donne accès à l’objet [Error](/javascript/api/office/office.error) .</span><span class="sxs-lookup"><span data-stu-id="c375b-123">If the call fails, the [error](/javascript/api/office/office.asyncresult#asynccontext) property of the `AsyncResult` object provides access to the [Error](/javascript/api/office/office.error) object.</span></span> <span data-ttu-id="c375b-124">Vous pouvez vérifier la valeur des propriétés [Error.name](/javascript/api/office/office.error#name) et [Error.message](/javascript/api/office/office.error#message) pour déterminer les raisons de l’échec de l’opération.</span><span class="sxs-lookup"><span data-stu-id="c375b-124">You can check the value of the [Error.name](/javascript/api/office/office.error#name) and [Error.message](/javascript/api/office/office.error#message) properties to determine why the set operation failed.</span></span> <span data-ttu-id="c375b-125">Sinon, le texte sélectionné dans le document s’affiche.</span><span class="sxs-lookup"><span data-stu-id="c375b-125">Otherwise, the selected text in the document is displayed.</span></span>

<span data-ttu-id="c375b-126">La propriété [AsyncResult.status](/javascript/api/office/office.asyncresult#error) est utilisée dans l’instruction **if** pour tester la réussite de l’appel.</span><span class="sxs-lookup"><span data-stu-id="c375b-126">The [AsyncResult.status](/javascript/api/office/office.asyncresult#error) property is used in the **if** statement to test whether the call succeeded.</span></span> <span data-ttu-id="c375b-127">[Office. AsyncResultStatus](/javascript/api/office/office.asyncresult#status) est une énumération des `AsyncResult.status` valeurs de propriété disponibles.</span><span class="sxs-lookup"><span data-stu-id="c375b-127">[Office.AsyncResultStatus](/javascript/api/office/office.asyncresult#status) is an enumeration of available `AsyncResult.status` property values.</span></span> <span data-ttu-id="c375b-128">`Office.AsyncResultStatus.Failed`donne la chaîne « failed » (et, à nouveau, peut également être spécifié comme cette chaîne littérale).</span><span class="sxs-lookup"><span data-stu-id="c375b-128">`Office.AsyncResultStatus.Failed` evaluates to the string "failed" (and, again, can also be specified as that literal string).</span></span>


## <a name="write-data-to-the-selection"></a><span data-ttu-id="c375b-129">Écriture de données dans la sélection</span><span class="sxs-lookup"><span data-stu-id="c375b-129">Write data to the selection</span></span>


<span data-ttu-id="c375b-130">L’exemple suivant montre comment définir la sélection pour afficher « Hello World! ».</span><span class="sxs-lookup"><span data-stu-id="c375b-130">The following example shows how to set the selection to show "Hello World!".</span></span>


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

<span data-ttu-id="c375b-p107">Le passage de différents types d’objets pour le paramètre  _data_ produit différents résultats. Le résultat varie en fonction de la sélection actuelle dans le document, de l’application qui héberge votre complément, et de l’éventuel passage forcé des données dans la sélection actuelle.</span><span class="sxs-lookup"><span data-stu-id="c375b-p107">Passing in different object types for the  _data_ parameter will have different results. The result depends on what is currently selected in the document, which application is hosting your add-in, and whether the data passed in can be coerced to the current selection.</span></span>

<span data-ttu-id="c375b-133">La fonction anonyme transmise dans la méthode [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) comme paramètre _callback_ est exécutée quand l’appel anonyme est terminé.</span><span class="sxs-lookup"><span data-stu-id="c375b-133">The anonymous function passed into the [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method as the _callback_ parameter is executed when the asynchronous call is completed.</span></span> <span data-ttu-id="c375b-134">Lorsque vous écrivez des données dans la sélection à l' `setSelectedDataAsync` aide de la méthode, le paramètre _asyncResult_ du rappel donne accès uniquement à l’état de l’appel et à l’objet d' [erreur](/javascript/api/office/office.error) en cas d’échec de l’appel.</span><span class="sxs-lookup"><span data-stu-id="c375b-134">When you write data to the selection by using the `setSelectedDataAsync` method, the _asyncResult_ parameter of the callback provides access only to the status of the call, and to the [Error](/javascript/api/office/office.error) object if the call fails.</span></span>

> [!NOTE]
> <span data-ttu-id="c375b-135">Depuis la publication d’Excel 2013 SP1 et de la version correspondante d’Excel sur le web, vous pouvez désormais [définir la mise en forme lors de l’écriture d’un tableau sur la sélection active](../excel/excel-add-ins-tables.md).</span><span class="sxs-lookup"><span data-stu-id="c375b-135">Starting with the release of the Excel 2013 SP1 and the corresponding build of Excel on the web, you can now [set formatting when writing a table to the current selection](../excel/excel-add-ins-tables.md).</span></span>


## <a name="detect-changes-in-the-selection"></a><span data-ttu-id="c375b-136">Détection de modifications dans la sélection</span><span class="sxs-lookup"><span data-stu-id="c375b-136">Detect changes in the selection</span></span>


<span data-ttu-id="c375b-137">L’exemple suivant montre comment détecter des modifications dans la sélection à l’aide de la méthode [Document.addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) permettant d’ajouter un gestionnaire d’événements pour l’événement [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) sur le document.</span><span class="sxs-lookup"><span data-stu-id="c375b-137">The following example shows how to detect changes in the selection by using the [Document.addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) method to add an event handler for the [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) event on the document.</span></span>


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

<span data-ttu-id="c375b-138">Le premier paramètre  _eventType_ spécifie le nom de l’événement auquel souscrire.</span><span class="sxs-lookup"><span data-stu-id="c375b-138">The first  _eventType_ parameter specifies the name of the event to subscribe to.</span></span> <span data-ttu-id="c375b-139">Le passage de `"documentSelectionChanged"` la chaîne pour ce paramètre équivaut à la `Office.EventType.DocumentSelectionChanged` transmission du type d’événement de l’énumération [Office. EventType](/javascript/api/office/office.eventtype) .</span><span class="sxs-lookup"><span data-stu-id="c375b-139">Passing the string `"documentSelectionChanged"` for this parameter is equivalent to passing the `Office.EventType.DocumentSelectionChanged` event type of the [Office.EventType](/javascript/api/office/office.eventtype) enumeration.</span></span>

<span data-ttu-id="c375b-p110">La fonction `myHander()` transmise dans la fonction comme deuxième paramètre _handler_ est un gestionnaire d’événements qui est exécuté lorsque la sélection change dans le document. La fonction est appelée avec un seul paramètre, _eventArgs_, qui contient une référence à un objet [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs) quand l’opération asynchrone se termine. Vous pouvez utiliser la propriété [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) pour accéder au document qui a déclenché l’événement.</span><span class="sxs-lookup"><span data-stu-id="c375b-p110">The  `myHander()` function that is passed into the function as the second _handler_ parameter is an event handler that is executed when the selection is changed on the document. The function is called with a single parameter, _eventArgs_, which will contain a reference to a [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs) object when the asynchronous operation completes. You can use the [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) property to access the document that raised the event.</span></span>


> [!NOTE]
> <span data-ttu-id="c375b-143">Vous pouvez ajouter plusieurs gestionnaires d’événements pour un événement donné en appelant à `addHandlerAsync` nouveau la méthode et en transmettant une fonction de gestionnaire d’événements supplémentaire pour le paramètre _handler_ .</span><span class="sxs-lookup"><span data-stu-id="c375b-143">You can add multiple event handlers for a given event by calling the `addHandlerAsync` method again and passing in an additional event handler function for the _handler_ parameter.</span></span> <span data-ttu-id="c375b-144">Cela fonctionnera correctement à condition que le nom de chaque fonction de gestionnaire d’événements soit unique.</span><span class="sxs-lookup"><span data-stu-id="c375b-144">This will work correctly as long as the name of each event handler function is unique.</span></span>


## <a name="stop-detecting-changes-in-the-selection"></a><span data-ttu-id="c375b-145">Arrêt de la détection de modifications dans la sélection</span><span class="sxs-lookup"><span data-stu-id="c375b-145">Stop detecting changes in the selection</span></span>


<span data-ttu-id="c375b-146">L’exemple suivant montre comment arrêter l’écoute de l’événement [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) en appelant la méthode [document.removeHandlerAsync](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="c375b-146">The following example shows how to stop listening to the [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) event by calling the [document.removeHandlerAsync](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-) method.</span></span>


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

<span data-ttu-id="c375b-147">Le `myHandler` nom de la fonction passé en tant que deuxième paramètre _handler_ désigne le gestionnaire d’événements qui sera supprimé de `SelectionChanged` l’événement.</span><span class="sxs-lookup"><span data-stu-id="c375b-147">The  `myHandler` function name that is passed as the second _handler_ parameter specifies the event handler that will be removed from the `SelectionChanged` event.</span></span>


> [!IMPORTANT]
> <span data-ttu-id="c375b-148">Si le paramètre facultatif _handler_ est omis lors de l' `removeHandlerAsync` appel de la méthode, tous les gestionnaires d’événements pour le paramètre _eventType_ spécifié sont supprimés.</span><span class="sxs-lookup"><span data-stu-id="c375b-148">If the optional  _handler_ parameter is omitted when the `removeHandlerAsync` method is called, all event handlers for the specified _eventType_ will be removed.</span></span>
