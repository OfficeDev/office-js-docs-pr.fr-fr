---
title: Utilisation d’événements à l’aide de l’API JavaScript pour Excel
description: ''
ms.date: 04/03/2019
localization_priority: Priority
ms.openlocfilehash: 7f05263f5220c2d60d0cebcfc686e1fed3f07900
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449266"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="86fb4-102">Utilisation d’événements à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="86fb4-102">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="86fb4-103">Cet article décrit des concepts importants relatifs à l’utilisation des événements dans Excel et fournit des exemples de code montrant comment inscrire des gestionnaires d’événements, gérer des événements et supprimer des gestionnaires d’événements à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="86fb4-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="86fb4-104">Événements dans Excel</span><span class="sxs-lookup"><span data-stu-id="86fb4-104">Events in Excel</span></span>

<span data-ttu-id="86fb4-p101">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span><span class="sxs-lookup"><span data-stu-id="86fb4-p101">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="86fb4-108">Événement</span><span class="sxs-lookup"><span data-stu-id="86fb4-108">Event</span></span> | <span data-ttu-id="86fb4-109">Description</span><span class="sxs-lookup"><span data-stu-id="86fb4-109">Description</span></span> | <span data-ttu-id="86fb4-110">Objets pris en charge</span><span class="sxs-lookup"><span data-stu-id="86fb4-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="86fb4-111">Se produit lorsqu’un objet est activé.</span><span class="sxs-lookup"><span data-stu-id="86fb4-111">Occurs when an object is activated.</span></span> | <span data-ttu-id="86fb4-112">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="86fb4-112">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onAdded` | <span data-ttu-id="86fb4-113">Se produit lorsqu’un objet est ajouté.</span><span class="sxs-lookup"><span data-stu-id="86fb4-113">Occurs when an object is added.</span></span> | <span data-ttu-id="86fb4-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="86fb4-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onCalculated` | <span data-ttu-id="86fb4-115">Se produit lorsqu’une feuille de calcul a terminé un calcul (ou que toutes les feuilles de calcul de la collection ont terminé).</span><span class="sxs-lookup"><span data-stu-id="86fb4-115">Occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="86fb4-116">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="86fb4-116">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="86fb4-117">Se produit lorsque les données au sein des cellules sont modifiées.</span><span class="sxs-lookup"><span data-stu-id="86fb4-117">Occurs when data within cells is changed.</span></span> | <span data-ttu-id="86fb4-118">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="86fb4-118">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onDataChanged` | <span data-ttu-id="86fb4-119">Se produit lors de la modification des données ou de la mise en forme dans la liaison.</span><span class="sxs-lookup"><span data-stu-id="86fb4-119">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="86fb4-120">**Binding**</span><span class="sxs-lookup"><span data-stu-id="86fb4-120">**Binding**</span></span>](/javascript/api/excel/excel.binding) |
| `onDeactivated` | <span data-ttu-id="86fb4-121">Se produit lorsqu’un objet est désactivé.</span><span class="sxs-lookup"><span data-stu-id="86fb4-121">Occurs when an object is deactivated.</span></span> | <span data-ttu-id="86fb4-122">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="86fb4-122">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDeleted` | <span data-ttu-id="86fb4-123">Se produit lorsqu’un objet est supprimé.</span><span class="sxs-lookup"><span data-stu-id="86fb4-123">Occurs when an object is deleted.</span></span> | <span data-ttu-id="86fb4-124">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="86fb4-124">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="86fb4-125">Se produit lorsque la cellule active ou la plage sélectionnée est modifiée.</span><span class="sxs-lookup"><span data-stu-id="86fb4-125">Occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="86fb4-126">[**Liaison**](/javascript/api/excel/excel.binding), [**Tableau**](/javascript/api/excel/excel.table),  [**Feuille de calcul**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="86fb4-126">[**Binding**](/javascript/api/excel/excel.binding), [**Table**](/javascript/api/excel/excel.table),  [**Worksheet**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="86fb4-127">Se produit lorsque les paramètres dans le document sont modifiés.</span><span class="sxs-lookup"><span data-stu-id="86fb4-127">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="86fb4-128">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="86fb4-128">**SettingCollection**</span></span>](/javascript/api/excel/excel.settingcollection) |

### <a name="events-in-preview"></a><span data-ttu-id="86fb4-129">Événements en préversion</span><span class="sxs-lookup"><span data-stu-id="86fb4-129">Events in preview</span></span>

> [!NOTE]
> <span data-ttu-id="86fb4-130">Les événements suivants sont actuellement disponibles uniquement en préversion publique.</span><span class="sxs-lookup"><span data-stu-id="86fb4-130">The following events are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| <span data-ttu-id="86fb4-131">Événement</span><span class="sxs-lookup"><span data-stu-id="86fb4-131">Event</span></span> | <span data-ttu-id="86fb4-132">Description</span><span class="sxs-lookup"><span data-stu-id="86fb4-132">Description</span></span> | <span data-ttu-id="86fb4-133">Objets pris en charge</span><span class="sxs-lookup"><span data-stu-id="86fb4-133">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="86fb4-134">Se produit lorsque la forme est activée.</span><span class="sxs-lookup"><span data-stu-id="86fb4-134">Occurs when the shape is activated.</span></span> | [<span data-ttu-id="86fb4-135">**Shape**</span><span class="sxs-lookup"><span data-stu-id="86fb4-135">**Shape**</span></span>](/javascript/api/excel/excel.shape)|
| `onAdded` | <span data-ttu-id="86fb4-136">Se produit lorsque le nouveau tableau est ajouté dans un classeur.</span><span class="sxs-lookup"><span data-stu-id="86fb4-136">Occurs when new table is added in a workbook.</span></span> | [<span data-ttu-id="86fb4-137">**TableCollection**</span><span class="sxs-lookup"><span data-stu-id="86fb4-137">**TableCollection**</span></span>](/javascript/api/excel/excel.tablecollection)|
| `onAutoSaveSettingChanged` | <span data-ttu-id="86fb4-138">Se produit lorsque le paramètre de `autoSave` est modifié dans le classeur.</span><span class="sxs-lookup"><span data-stu-id="86fb4-138">Occurs when the `autoSave` setting is changed on the workbook.</span></span> | [<span data-ttu-id="86fb4-139">**Classeur**</span><span class="sxs-lookup"><span data-stu-id="86fb4-139">**Workbook**</span></span>](/javascript/api/excel/excel.workbook) |
| `onChanged` | <span data-ttu-id="86fb4-140">Se produit lorsqu’une feuille de calcul dans le classeur est modifiée.</span><span class="sxs-lookup"><span data-stu-id="86fb4-140">Occurs when any worksheet in the workbook is changed.</span></span> | [<span data-ttu-id="86fb4-141">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="86fb4-141">**WorksheetCollection**</span></span>](/javascript/api/excel/excel.worksheetcollection)|
| `onDeactivated` | <span data-ttu-id="86fb4-142">Se produit lorsque la forme est désactivée.</span><span class="sxs-lookup"><span data-stu-id="86fb4-142">Occurs when the shape is deactivated.</span></span> | [<span data-ttu-id="86fb4-143">**Shape**</span><span class="sxs-lookup"><span data-stu-id="86fb4-143">**Shape**</span></span>](/javascript/api/excel/excel.shape)|
| `onDeleted` | <span data-ttu-id="86fb4-144">Se produit lorsque le tableau spécifié est supprimé dans un classeur.</span><span class="sxs-lookup"><span data-stu-id="86fb4-144">Occurs when the specified table is deleted in a workbook.</span></span> | [<span data-ttu-id="86fb4-145">**TableCollection**</span><span class="sxs-lookup"><span data-stu-id="86fb4-145">**TableCollection**</span></span>](/javascript/api/excel/excel.tablecollection)|
| `onFiltered` | <span data-ttu-id="86fb4-146">Se produit lorsque le filtre est appliqué sur un objet.</span><span class="sxs-lookup"><span data-stu-id="86fb4-146">Occurs when filter is applied on an object.</span></span> | <span data-ttu-id="86fb4-147">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="86fb4-147">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onFormatChanged` | <span data-ttu-id="86fb4-148">Se produit lorsque le format est modifié sur une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="86fb4-148">Occurs when the format is changed on a worksheet.</span></span> | <span data-ttu-id="86fb4-149">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="86fb4-149">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="86fb4-150">Se produit lorsque la sélection change sur n’importe quelle feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="86fb4-150">Occurs when the selection changes on any worksheet.</span></span> | [<span data-ttu-id="86fb4-151">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="86fb4-151">**WorksheetCollection**</span></span>](/javascript/api/excel/excel.worksheetcollection) |

### <a name="event-triggers"></a><span data-ttu-id="86fb4-152">Déclencheurs d’événements</span><span class="sxs-lookup"><span data-stu-id="86fb4-152">Event triggers</span></span>

<span data-ttu-id="86fb4-153">Événements au sein d’un classeur Excel pouvant être déclenchés par :</span><span class="sxs-lookup"><span data-stu-id="86fb4-153">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="86fb4-154">Interaction de l’utilisateur via l’interface utilisateur Excel (IU) modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="86fb4-154">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="86fb4-155">Complément (JavaScript) Office modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="86fb4-155">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="86fb4-156">Complément VBA (macro) modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="86fb4-156">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="86fb4-157">Toute modification conforme aux comportements par défaut d’Excel déclenche les événements correspondants dans un classeur.</span><span class="sxs-lookup"><span data-stu-id="86fb4-157">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="86fb4-158">Cycle de vie d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="86fb4-158">Lifecycle of an event handler</span></span>

<span data-ttu-id="86fb4-159">Un gestionnaire d’événements est créé lorsqu’un complément inscrit le gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="86fb4-159">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="86fb4-160">Il est détruit lorsque le complément annule l’inscription du gestionnaire d’événements ou lorsque le complément est actualisé, rechargé ou fermé.</span><span class="sxs-lookup"><span data-stu-id="86fb4-160">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="86fb4-161">Les gestionnaires d’événements ne sont pas conservés dans le fichier Excel ou entre des sessions avec Excel Online.</span><span class="sxs-lookup"><span data-stu-id="86fb4-161">Event handlers do not persist as part of the Excel file, or across sessions with Excel Online.</span></span>

> [!CAUTION]
> <span data-ttu-id="86fb4-162">Lorsqu’un objet dans lequel des événements sont inscrits est supprimé (par exemple, un tableau avec un événement `onChanged`), le gestionnaire d’événements n’est plus déclenché mais reste en mémoire jusqu’à ce que le complément ou la session Excel soit actualisé(e) ou se ferme.</span><span class="sxs-lookup"><span data-stu-id="86fb4-162">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="86fb4-163">Événements et co-création</span><span class="sxs-lookup"><span data-stu-id="86fb4-163">Events and coauthoring</span></span>

<span data-ttu-id="86fb4-p104">Avec la [co-création](co-authoring-in-excel-add-ins.md), plusieurs personnes peuvent travailler ensemble et modifier le même classeur Excel simultanément. Pour les événements pouvant être déclenchés par un co-auteur, tels que `onChanged`, l’objet **Event** correspondant contient une propriété **source** qui indique si l’événement a été déclenché localement par l’utilisateur actuel (`event.source = Local`) ou par le co-auteur à distance (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="86fb4-p104">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="86fb4-166">Inscription d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="86fb4-166">Register an event handler</span></span>

<span data-ttu-id="86fb4-p105">L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onChanged` dans la feuille de calcul **Sample**. Le code indique que la fonction `handleDataChange` doit être exécutée lorsque les données de la feuille de calcul sont modifiées.</span><span class="sxs-lookup"><span data-stu-id="86fb4-p105">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="86fb4-169">Gestion d’un événement</span><span class="sxs-lookup"><span data-stu-id="86fb4-169">Handle an event</span></span>

<span data-ttu-id="86fb4-p106">Comme indiqué dans l’exemple précédent, lorsque vous inscrivez un gestionnaire d’événements, vous indiquez la fonction devant être exécutée lorsque l’événement spécifié se produit. Vous pouvez créer cette fonction pour effectuer n’importe quelle action nécessaire à votre scénario. L’exemple de code suivant montre une fonction de gestionnaire d’événements qui écrit simplement des informations sur l’événement dans la console.</span><span class="sxs-lookup"><span data-stu-id="86fb4-p106">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span>

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="86fb4-173">Suppression d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="86fb4-173">Remove an event handler</span></span>

<span data-ttu-id="86fb4-p107">L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onSelectionChanged` dans la feuille de calcul **Sample** et définit la fonction `handleSelectionChange` qui est exécutée lorsqu’un événement se produit. Il définit également la fonction `remove()` pouvant être appelée par la suite pour supprimer ce gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="86fb4-p107">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="86fb4-176">Activation et désactivation d’événements</span><span class="sxs-lookup"><span data-stu-id="86fb4-176">Enable and disable events</span></span>

<span data-ttu-id="86fb4-177">La performance d’un complément peut être améliorée en désactivant les événements.</span><span class="sxs-lookup"><span data-stu-id="86fb4-177">The performance of an add-in may be improved by disabling events.</span></span>
<span data-ttu-id="86fb4-178">Par exemple, il se peut que votre application ne doive jamais recevoir d’événements, ou elle peut ignorer des événements lors de modifications par lots de plusieurs entités.</span><span class="sxs-lookup"><span data-stu-id="86fb4-178">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="86fb4-179">Les événements sont activés et désactivés au niveau [runtime](/javascript/api/excel/excel.runtime).</span><span class="sxs-lookup"><span data-stu-id="86fb4-179">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="86fb4-180">La propriété `enableEvents` détermine si les événements sont déclenchés et leurs gestionnaires activés.</span><span class="sxs-lookup"><span data-stu-id="86fb4-180">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="86fb4-181">L’exemple de code suivant montre comment activer et désactiver des événements.</span><span class="sxs-lookup"><span data-stu-id="86fb4-181">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="86fb4-182">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="86fb4-182">See also</span></span>

- [<span data-ttu-id="86fb4-183">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="86fb4-183">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
