---
title: Utilisation d’événements à l’aide de l’API JavaScript pour Excel
description: ''
ms.date: 09/21/2018
ms.openlocfilehash: b56d25e7e0306b4881115397d4136e63ddc03e5c
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459174"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="98cef-102">Utilisation d’événements à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="98cef-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="98cef-103">Cet article décrit des concepts importants relatifs à l’utilisation des événements dans Excel et fournit des exemples de code montrant comment inscrire des gestionnaires d’événements, gérer des événements et supprimer des gestionnaires d’événements à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="98cef-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="98cef-104">Événements dans Excel</span><span class="sxs-lookup"><span data-stu-id="98cef-104">Events in Excel</span></span>

<span data-ttu-id="98cef-p101">Chaque fois que certains types de modifications se produisent dans un classeur Excel, une notification d’événement se déclenche. En utilisant l’API JavaScript pour Excel, vous pouvez inscrire les gestionnaires d’événements autorisant votre complément à exécuter automatiquement une fonction désignée lorsqu’un événement spécifique se produit. Les événements suivants sont actuellement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="98cef-p101">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="98cef-108">Événement</span><span class="sxs-lookup"><span data-stu-id="98cef-108">Event</span></span> | <span data-ttu-id="98cef-109">Description</span><span class="sxs-lookup"><span data-stu-id="98cef-109">Description</span></span> | <span data-ttu-id="98cef-110">Objets pris en charge</span><span class="sxs-lookup"><span data-stu-id="98cef-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="98cef-111">Événement se produisant lors de l’ajout d’un objet.</span><span class="sxs-lookup"><span data-stu-id="98cef-111">Event that occurs when an object is added.</span></span> | <span data-ttu-id="98cef-112">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="98cef-112">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDeleted` | <span data-ttu-id="98cef-113">Événement se produisant lorsqu’un objet est supprimé.</span><span class="sxs-lookup"><span data-stu-id="98cef-113">Event that occurs when an object is deleted.</span></span> | <span data-ttu-id="98cef-114">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="98cef-114">[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onActivated` | <span data-ttu-id="98cef-115">Événement se produisant lorsqu’un objet est activé.</span><span class="sxs-lookup"><span data-stu-id="98cef-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="98cef-116">[**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="98cef-116">[**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span></span> |
| `onDeactivated` | <span data-ttu-id="98cef-117">Événement se produisant lorsqu’un objet est désactivé.</span><span class="sxs-lookup"><span data-stu-id="98cef-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="98cef-118">[**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="98cef-118">[**Chart**](https://docs.microsoft.com/javascript/api/excel/excel.chart), [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span></span> |
| `onCalculated` | <span data-ttu-id="98cef-119">Événement qui se produit lorsqu'une feuille de calcul a terminé le calcul (ou que toutes les feuilles de calcul de la collection sont terminées).</span><span class="sxs-lookup"><span data-stu-id="98cef-119">Event that occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="98cef-120">[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="98cef-120">**WorksheetCollection**, [Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="98cef-121">Événement se produisant lorsque les données au sein des cellules sont modifiées.</span><span class="sxs-lookup"><span data-stu-id="98cef-121">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="98cef-122">[**Feuille de calcul**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), [**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table), [**TableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection)</span><span class="sxs-lookup"><span data-stu-id="98cef-122">[**Worksheet**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), [**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table), [**TableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="98cef-123">Événement se produisant lors de la modification des données ou de la mise en forme dans la liaison.</span><span class="sxs-lookup"><span data-stu-id="98cef-123">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="98cef-124">**Liaison**</span><span class="sxs-lookup"><span data-stu-id="98cef-124">**Binding**</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.binding) |
| `onSelectionChanged` | <span data-ttu-id="98cef-125">Événement se produisant lorsque la cellule active ou la plage sélectionnée est modifiée.</span><span class="sxs-lookup"><span data-stu-id="98cef-125">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="98cef-126">[**Feuille de calcul**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), [**Table**](https://docs.microsoft.com/javascript/api/excel/excel.table), [**Liaison**](https://docs.microsoft.com/javascript/api/excel/excel.binding)</span><span class="sxs-lookup"><span data-stu-id="98cef-126">**Worksheet**, [Table](https://docs.microsoft.com/javascript/api/excel/excel.worksheet), **Binding**</span></span> |
| `onSettingsChanged` | <span data-ttu-id="98cef-127">Événement qui se produit lorsque les Paramètres dans le document sont modifiés.</span><span class="sxs-lookup"><span data-stu-id="98cef-127">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="98cef-128">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="98cef-128">**SettingCollection**</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection) |

### <a name="event-triggers"></a><span data-ttu-id="98cef-129">Déclencheurs d’événements</span><span class="sxs-lookup"><span data-stu-id="98cef-129">Event triggers</span></span>

<span data-ttu-id="98cef-130">Événements au sein d’un classeur Excel pouvant être déclenchés par :</span><span class="sxs-lookup"><span data-stu-id="98cef-130">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="98cef-131">Interaction de l’utilisateur via l’interface utilisateur Excel (IU) modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="98cef-131">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="98cef-132">Complément (JavaScript) Office modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="98cef-132">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="98cef-133">Complément VBA (macro) modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="98cef-133">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="98cef-134">Toute modification conforme aux comportements par défaut d’Excel déclenche les événements correspondants dans un classeur.</span><span class="sxs-lookup"><span data-stu-id="98cef-134">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="98cef-135">Cycle de vie d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="98cef-135">Lifecycle of an event handler</span></span>

<span data-ttu-id="98cef-p102">Un gestionnaire d’événements est créé lorsqu’un complément inscrit le gestionnaire d’événements et est détruit lorsque le complément désinscrit le gestionnaire d’événements ou que le complément est fermé. Les gestionnaires d’événements ne persistent pas en tant que partie du fichier Excel.</span><span class="sxs-lookup"><span data-stu-id="98cef-p102">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="98cef-138">Événements et co-création</span><span class="sxs-lookup"><span data-stu-id="98cef-138">Events and coauthoring</span></span>

<span data-ttu-id="98cef-p103">Avec la [co-création](co-authoring-in-excel-add-ins.md), plusieurs personnes peuvent travailler ensemble et modifier le même classeur Excel simultanément. Pour les événements pouvant être déclenchés par un co-auteur, tels que `onChanged`, l’objet **Event** correspondant contient une propriété **source** qui indique si l’événement a été déclenché localement par l’utilisateur actuel (`event.source = Local`) ou par le co-auteur à distance (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="98cef-p103">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="98cef-141">Inscription d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="98cef-141">Register an event handler</span></span>

<span data-ttu-id="98cef-p104">L’exemple de code suivant enregistre un gestionnaire d’événements pour le `onChanged` événement dans la feuille de calcul nommée **Sample**. Le code spécifie que lors de la modification des données dans cette feuille de calcul, la `handleDataChange` fonction doit s’exécuter.</span><span class="sxs-lookup"><span data-stu-id="98cef-p104">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

```js
Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    worksheet.onChanged.add(handleChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);
```

## <a name="handle-an-event"></a><span data-ttu-id="98cef-144">Gestion d’un événement</span><span class="sxs-lookup"><span data-stu-id="98cef-144">Handle an event</span></span>

<span data-ttu-id="98cef-p105">Comme indiqué dans l’exemple précédent, lorsque vous inscrivez un gestionnaire d’événements, vous indiquez la fonction devant être exécutée lorsque l’événement spécifié se produit. Vous pouvez créer cette fonction pour effectuer n’importe quelle action nécessaire à votre scénario. L’exemple de code suivant montre une fonction de gestionnaire d’événements qui écrit simplement des informations sur l’événement dans la console.</span><span class="sxs-lookup"><span data-stu-id="98cef-p105">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

```js
function handleChange(event)
{ 
    return Excel.run(function(context){
        return context.sync()
            .then(function() {
                console.log("Change type of event: " + event.changeType);
                console.log("Address of event: " + event.address);
                console.log("Source of event: " + event.source);
            });
    }).catch(errorHandlerFunction);
}
```

## <a name="remove-an-event-handler"></a><span data-ttu-id="98cef-148">Suppression d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="98cef-148">Remove an event handler</span></span>

<span data-ttu-id="98cef-p106">L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onSelectionChanged` dans la feuille de calcul **Sample** et définit la fonction `handleSelectionChange` qui est exécutée lorsqu’un événement se produit. Il définit également la fonction `remove()` pouvant être appelée par la suite pour supprimer ce gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="98cef-p106">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

```js
var eventResult;

Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);

function handleSelectionChange(event)
{ 
    return Excel.run(function(context){
        return context.sync()
            .then(function() {
                console.log("Address of current selection: " + event.address);
            });
    }).catch(errorHandlerFunction);
}

function remove() {
    return Excel.run(eventResult.context, function (context) {
        eventResult.remove();
        
        return context.sync()
            .then(function() {
                eventResult = null;
                console.log("Event handler successfully removed.");
            });
    }).catch(errorHandlerFunction);
}
```

## <a name="enable-and-disable-events"></a><span data-ttu-id="98cef-151">Activer et désactiver des événements</span><span class="sxs-lookup"><span data-stu-id="98cef-151">Enable and disable events</span></span>

<span data-ttu-id="98cef-p107">Le niveau de performance d’un complément peut être amélioré en désactivant des événements. Par exemple, votre application pourrait ne jamais avoir besoin de recevoir des événements, ou bien elle pourrait ignorer les événements lors de l’exécution de lots de modifications de plusieurs entités.</span><span class="sxs-lookup"><span data-stu-id="98cef-p107">The performance of an add-in may be improved by disabling events. For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span> 

<span data-ttu-id="98cef-p108">Les événements sont activés et désactivés au niveau de [l’exécution](https://docs.microsoft.com/javascript/api/excel/excel.runtime). La propriété `enableEvents` détermine si les événements sont déclenchés et si leurs gestionnaires sont activés.</span><span class="sxs-lookup"><span data-stu-id="98cef-p108">Events are enabled and disabled at the [runtime](https://docs.microsoft.com/javascript/api/excel/excel.runtime) level. The `enableEvents` property determines if events are fired and their handlers are activated.</span></span> 

<span data-ttu-id="98cef-156">L’exemple de code suivant montre comment activer ou désactiver les événements.</span><span class="sxs-lookup"><span data-stu-id="98cef-156">The following code sample shows how to toggle events on and off.</span></span>

```js
Excel.run(function (context) {
    context.runtime.load("enableEvents");
    return context.sync()
        .then(function () {
            var eventBoolean = !context.runtime.enableEvents;
            context.runtime.enableEvents = eventBoolean;
            if (eventBoolean) {
                console.log("Events are currently on.");
            } else {
                console.log("Events are currently off.");
            }
        }).then(context.sync);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a><span data-ttu-id="98cef-157">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="98cef-157">See also</span></span>

- [<span data-ttu-id="98cef-158">Concepts fondamentaux de programmation avec l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="98cef-158">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)