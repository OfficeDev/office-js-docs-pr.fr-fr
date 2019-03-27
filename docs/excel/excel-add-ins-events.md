---
title: Utilisation d’événements à l’aide de l’API JavaScript pour Excel
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 08653a84c051709d16371d89672d3f7ebe2030b7
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872017"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="4e310-102">Utilisation d’événements à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="4e310-102">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="4e310-103">Cet article décrit des concepts importants relatifs à l’utilisation des événements dans Excel et fournit des exemples de code montrant comment inscrire des gestionnaires d’événements, gérer des événements et supprimer des gestionnaires d’événements à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="4e310-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="4e310-104">Événements dans Excel</span><span class="sxs-lookup"><span data-stu-id="4e310-104">Events in Excel</span></span>

<span data-ttu-id="4e310-p101">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span><span class="sxs-lookup"><span data-stu-id="4e310-p101">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="4e310-108">Événement</span><span class="sxs-lookup"><span data-stu-id="4e310-108">Event</span></span> | <span data-ttu-id="4e310-109">Description</span><span class="sxs-lookup"><span data-stu-id="4e310-109">Description</span></span> | <span data-ttu-id="4e310-110">Objets pris en charge</span><span class="sxs-lookup"><span data-stu-id="4e310-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="4e310-111">Événement se produisant lors de l’ajout d’un objet.</span><span class="sxs-lookup"><span data-stu-id="4e310-111">Event that occurs when an object is added.</span></span> | <span data-ttu-id="4e310-112">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="4e310-112">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDeleted` | <span data-ttu-id="4e310-113">Événement se produisant lorsqu’un objet est supprimé.</span><span class="sxs-lookup"><span data-stu-id="4e310-113">Event that occurs when an object is deleted.</span></span> | <span data-ttu-id="4e310-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="4e310-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onActivated` | <span data-ttu-id="4e310-115">Événement se produisant lorsqu’un objet est activé.</span><span class="sxs-lookup"><span data-stu-id="4e310-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="4e310-116">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="4e310-116">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onDeactivated` | <span data-ttu-id="4e310-117">Événement se produisant lorsqu’un objet est désactivé.</span><span class="sxs-lookup"><span data-stu-id="4e310-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="4e310-118">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="4e310-118">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onCalculated` | <span data-ttu-id="4e310-119">Événement se produisant lorsqu’une feuille de calcul a terminé un calcul (ou toutes les feuilles de calcul de la collection ont terminé).</span><span class="sxs-lookup"><span data-stu-id="4e310-119">Event that occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="4e310-120">[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="4e310-120">[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onChanged` | <span data-ttu-id="4e310-121">Événement se produisant lorsque les données au sein des cellules sont modifiées.</span><span class="sxs-lookup"><span data-stu-id="4e310-121">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="4e310-122">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection)</span><span class="sxs-lookup"><span data-stu-id="4e310-122">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="4e310-123">Événement se produisant lors de la modification des données ou de la mise en forme dans la liaison.</span><span class="sxs-lookup"><span data-stu-id="4e310-123">Event that occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="4e310-124">**Binding**</span><span class="sxs-lookup"><span data-stu-id="4e310-124">**Binding**</span></span>](/javascript/api/excel/excel.binding) |
| `onSelectionChanged` | <span data-ttu-id="4e310-125">Événement se produisant lorsque la cellule active ou la plage sélectionnée est modifiée.</span><span class="sxs-lookup"><span data-stu-id="4e310-125">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="4e310-126">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**Table**](/javascript/api/excel/excel.table), [**Binding**](/javascript/api/excel/excel.binding)</span><span class="sxs-lookup"><span data-stu-id="4e310-126">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**Table**](/javascript/api/excel/excel.table), [**Binding**](/javascript/api/excel/excel.binding)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="4e310-127">Événement se produisant lorsque les paramètres dans le document sont modifiés.</span><span class="sxs-lookup"><span data-stu-id="4e310-127">Event that occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="4e310-128">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="4e310-128">**SettingCollection**</span></span>](/javascript/api/excel/excel.settingcollection) |

### <a name="event-triggers"></a><span data-ttu-id="4e310-129">Déclencheurs d’événements</span><span class="sxs-lookup"><span data-stu-id="4e310-129">Event triggers</span></span>

<span data-ttu-id="4e310-130">Événements au sein d’un classeur Excel pouvant être déclenchés par :</span><span class="sxs-lookup"><span data-stu-id="4e310-130">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="4e310-131">Interaction de l’utilisateur via l’interface utilisateur Excel (IU) modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="4e310-131">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="4e310-132">Complément (JavaScript) Office modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="4e310-132">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="4e310-133">Complément VBA (macro) modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="4e310-133">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="4e310-134">Toute modification conforme aux comportements par défaut d’Excel déclenche les événements correspondants dans un classeur.</span><span class="sxs-lookup"><span data-stu-id="4e310-134">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="4e310-135">Cycle de vie d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="4e310-135">Lifecycle of an event handler</span></span>

<span data-ttu-id="4e310-136">Un gestionnaire d’événements est créé lorsqu’un complément inscrit le gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="4e310-136">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="4e310-137">Il est détruit lorsque le complément annule l’inscription du gestionnaire d’événements ou lorsque le complément est actualisé, rechargé ou fermé.</span><span class="sxs-lookup"><span data-stu-id="4e310-137">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="4e310-138">Les gestionnaires d’événements ne sont pas conservés dans le fichier Excel ou entre des sessions avec Excel Online.</span><span class="sxs-lookup"><span data-stu-id="4e310-138">Event handlers do not persist as part of the Excel file, or across sessions with Excel Online.</span></span>

> [!CAUTION]
> <span data-ttu-id="4e310-139">Lorsqu’un objet dans lequel des événements sont inscrits est supprimé (par exemple, un tableau avec un événement `onChanged`), le gestionnaire d’événements n’est plus déclenché mais reste en mémoire jusqu’à ce que le complément ou la session Excel soit actualisé(e) ou se ferme.</span><span class="sxs-lookup"><span data-stu-id="4e310-139">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="4e310-140">Événements et co-création</span><span class="sxs-lookup"><span data-stu-id="4e310-140">Events and coauthoring</span></span>

<span data-ttu-id="4e310-p103">Avec la [co-création](co-authoring-in-excel-add-ins.md), plusieurs personnes peuvent travailler ensemble et modifier le même classeur Excel simultanément. Pour les événements pouvant être déclenchés par un co-auteur, tels que `onChanged`, l’objet **Event** correspondant contient une propriété **source** qui indique si l’événement a été déclenché localement par l’utilisateur actuel (`event.source = Local`) ou par le co-auteur à distance (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="4e310-p103">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="4e310-143">Inscription d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="4e310-143">Register an event handler</span></span>

<span data-ttu-id="4e310-p104">L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onChanged` dans la feuille de calcul **Sample**. Le code indique que la fonction `handleDataChange` doit être exécutée lorsque les données de la feuille de calcul sont modifiées.</span><span class="sxs-lookup"><span data-stu-id="4e310-p104">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="4e310-146">Gestion d’un événement</span><span class="sxs-lookup"><span data-stu-id="4e310-146">Handle an event</span></span>

<span data-ttu-id="4e310-p105">Comme indiqué dans l’exemple précédent, lorsque vous inscrivez un gestionnaire d’événements, vous indiquez la fonction devant être exécutée lorsque l’événement spécifié se produit. Vous pouvez créer cette fonction pour effectuer n’importe quelle action nécessaire à votre scénario. L’exemple de code suivant montre une fonction de gestionnaire d’événements qui écrit simplement des informations sur l’événement dans la console.</span><span class="sxs-lookup"><span data-stu-id="4e310-p105">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="4e310-150">Suppression d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="4e310-150">Remove an event handler</span></span>

<span data-ttu-id="4e310-p106">L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onSelectionChanged` dans la feuille de calcul **Sample** et définit la fonction `handleSelectionChange` qui est exécutée lorsqu’un événement se produit. Il définit également la fonction `remove()` pouvant être appelée par la suite pour supprimer ce gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="4e310-p106">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="4e310-153">Activation et désactivation d’événements</span><span class="sxs-lookup"><span data-stu-id="4e310-153">Enable and disable events</span></span>

<span data-ttu-id="4e310-154">La performance d’un complément peut être améliorée en désactivant les événements.</span><span class="sxs-lookup"><span data-stu-id="4e310-154">The performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="4e310-155">Par exemple, il se peut que votre application ne doive jamais recevoir d’événements, ou elle peut ignorer des événements lors de modifications par lots de plusieurs entités.</span><span class="sxs-lookup"><span data-stu-id="4e310-155">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="4e310-156">Les événements sont activés et désactivés au niveau [runtime](/javascript/api/excel/excel.runtime).</span><span class="sxs-lookup"><span data-stu-id="4e310-156">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="4e310-157">La propriété `enableEvents` détermine si les événements sont déclenchés et leurs gestionnaires activés.</span><span class="sxs-lookup"><span data-stu-id="4e310-157">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="4e310-158">L’exemple de code suivant montre comment activer et désactiver des événements.</span><span class="sxs-lookup"><span data-stu-id="4e310-158">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="4e310-159">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="4e310-159">See also</span></span>

- [<span data-ttu-id="4e310-160">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="4e310-160">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
