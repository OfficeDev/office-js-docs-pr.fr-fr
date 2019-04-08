---
title: Utilisation d’événements à l’aide de l’API JavaScript pour Excel
description: ''
ms.date: 04/03/2019
localization_priority: Priority
ms.openlocfilehash: 7f05263f5220c2d60d0cebcfc686e1fed3f07900
ms.sourcegitcommit: 63219bcc1bb5e3bed7eb6c6b0adb73a4829c7e8f
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/05/2019
ms.locfileid: "31479710"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="bf436-102">Utilisation d’événements à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="bf436-102">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="bf436-103">Cet article décrit des concepts importants relatifs à l’utilisation des événements dans Excel et fournit des exemples de code montrant comment inscrire des gestionnaires d’événements, gérer des événements et supprimer des gestionnaires d’événements à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="bf436-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="bf436-104">Événements dans Excel</span><span class="sxs-lookup"><span data-stu-id="bf436-104">Events in Excel</span></span>

<span data-ttu-id="bf436-p101">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span><span class="sxs-lookup"><span data-stu-id="bf436-p101">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="bf436-108">Événement</span><span class="sxs-lookup"><span data-stu-id="bf436-108">Event</span></span> | <span data-ttu-id="bf436-109">Description</span><span class="sxs-lookup"><span data-stu-id="bf436-109">Description</span></span> | <span data-ttu-id="bf436-110">Objets pris en charge</span><span class="sxs-lookup"><span data-stu-id="bf436-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="bf436-111">Se produit lorsqu’un objet est activé.</span><span class="sxs-lookup"><span data-stu-id="bf436-111">Event that occurs when an object is added.</span></span> | <span data-ttu-id="bf436-112">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="bf436-112">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheet), [**Worksheet**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onAdded` | <span data-ttu-id="bf436-113">Se produit lorsqu’un objet est ajouté.</span><span class="sxs-lookup"><span data-stu-id="bf436-113">Event that occurs when an object is deleted.</span></span> | <span data-ttu-id="bf436-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="bf436-114">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onCalculated` | <span data-ttu-id="bf436-115">Se produit lorsqu’une feuille de calcul a terminé un calcul (ou que toutes les feuilles de calcul de la collection ont terminé).</span><span class="sxs-lookup"><span data-stu-id="bf436-115">Event that occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="bf436-116">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="bf436-116">[**WorksheetCollection**](/javascript/api/excel/excel.worksheet), [**Worksheet**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="bf436-117">Se produit lorsque les données au sein des cellules sont modifiées.</span><span class="sxs-lookup"><span data-stu-id="bf436-117">Occurs when data within the binding is changed.</span></span> | <span data-ttu-id="bf436-118">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="bf436-118">[**Worksheet**](/javascript/api/excel/excel.table), [**Table**](/javascript/api/excel/excel.tablecollection), [**TableCollection**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onDataChanged` | <span data-ttu-id="bf436-119">Se produit lors de la modification des données ou de la mise en forme dans la liaison.</span><span class="sxs-lookup"><span data-stu-id="bf436-119">Occurs when data or formatting within the binding is changed.</span></span> | [**<span data-ttu-id="bf436-120">Liaison</span><span class="sxs-lookup"><span data-stu-id="bf436-120">Binding</span></span>**](/javascript/api/excel/excel.binding) |
| `onDeactivated` | <span data-ttu-id="bf436-121">Se produit lorsqu’un objet est désactivé.</span><span class="sxs-lookup"><span data-stu-id="bf436-121">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="bf436-122">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="bf436-122">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheet), [**Worksheet**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDeleted` | <span data-ttu-id="bf436-123">Se produit lorsqu’un objet est supprimé.</span><span class="sxs-lookup"><span data-stu-id="bf436-123">Event that occurs when an object is deleted.</span></span> | <span data-ttu-id="bf436-124">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="bf436-124">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="bf436-125">Se produit lorsque la cellule active ou la plage sélectionnée est modifiée.</span><span class="sxs-lookup"><span data-stu-id="bf436-125">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="bf436-126">[**Liaison**](/javascript/api/excel/excel.binding), [**Tableau**](/javascript/api/excel/excel.table),  [**Feuille de calcul**](/javascript/api/excel/excel.worksheet)</span><span class="sxs-lookup"><span data-stu-id="bf436-126">[**Binding**](/javascript/api/excel/excel.binding), [**Table**](/javascript/api/excel/excel.table),  [**Worksheet**](/javascript/api/excel/excel.worksheet)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="bf436-127">Se produit lorsque les paramètres dans le document sont modifiés.</span><span class="sxs-lookup"><span data-stu-id="bf436-127">Occurs when the Settings in the document are changed.</span></span> | [**<span data-ttu-id="bf436-128">SettingCollection</span><span class="sxs-lookup"><span data-stu-id="bf436-128">SettingCollection</span></span>**](/javascript/api/excel/excel.settingcollection) |

### <a name="events-in-preview"></a><span data-ttu-id="bf436-129">Événements en préversion</span><span class="sxs-lookup"><span data-stu-id="bf436-129">Events in preview</span></span>

> [!NOTE]
> <span data-ttu-id="bf436-130">Les événements suivants sont actuellement disponibles uniquement en préversion publique.</span><span class="sxs-lookup"><span data-stu-id="bf436-130">The  and  methods are currently available only in public preview (beta).</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| <span data-ttu-id="bf436-131">Événement</span><span class="sxs-lookup"><span data-stu-id="bf436-131">Event</span></span> | <span data-ttu-id="bf436-132">Description</span><span class="sxs-lookup"><span data-stu-id="bf436-132">Description</span></span> | <span data-ttu-id="bf436-133">Objets pris en charge</span><span class="sxs-lookup"><span data-stu-id="bf436-133">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="bf436-134">Se produit lorsque la forme est activée.</span><span class="sxs-lookup"><span data-stu-id="bf436-134">Occurs when the shape is activated.</span></span> | [**<span data-ttu-id="bf436-135">Forme</span><span class="sxs-lookup"><span data-stu-id="bf436-135">Shape</span></span>**](/javascript/api/excel/excel.shape)|
| `onAdded` | <span data-ttu-id="bf436-136">Se produit lorsque le nouveau tableau est ajouté dans un classeur.</span><span class="sxs-lookup"><span data-stu-id="bf436-136">Occurs when new table is added in a workbook.</span></span> | [**<span data-ttu-id="bf436-137">TableCollection</span><span class="sxs-lookup"><span data-stu-id="bf436-137">tableCollection</span></span>**](/javascript/api/excel/excel.tablecollection)|
| `onAutoSaveSettingChanged` | <span data-ttu-id="bf436-138">Se produit lorsque le paramètre de `autoSave` est modifié dans le classeur.</span><span class="sxs-lookup"><span data-stu-id="bf436-138">Occurs when the autoSave setting is changed on the workbook.</span></span> | [**<span data-ttu-id="bf436-139">Classeur</span><span class="sxs-lookup"><span data-stu-id="bf436-139">Workbook</span></span>**](/javascript/api/excel/excel.workbook) |
| `onChanged` | <span data-ttu-id="bf436-140">Se produit lorsqu’une feuille de calcul dans le classeur est modifiée.</span><span class="sxs-lookup"><span data-stu-id="bf436-140">Occurs when any worksheet in the workbook is changed.</span></span> | [**<span data-ttu-id="bf436-141">WorksheetCollection</span><span class="sxs-lookup"><span data-stu-id="bf436-141">worksheetCollection</span></span>**](/javascript/api/excel/excel.worksheetcollection)|
| `onDeactivated` | <span data-ttu-id="bf436-142">Se produit lorsque la forme est désactivée.</span><span class="sxs-lookup"><span data-stu-id="bf436-142">Occurs when the shape is deactivated.</span></span> | [**<span data-ttu-id="bf436-143">Forme</span><span class="sxs-lookup"><span data-stu-id="bf436-143">Shape</span></span>**](/javascript/api/excel/excel.shape)|
| `onDeleted` | <span data-ttu-id="bf436-144">Se produit lorsque le tableau spécifié est supprimé dans un classeur.</span><span class="sxs-lookup"><span data-stu-id="bf436-144">Occurs when the specified table is deleted in a workbook.</span></span> | [**<span data-ttu-id="bf436-145">TableCollection</span><span class="sxs-lookup"><span data-stu-id="bf436-145">tableCollection</span></span>**](/javascript/api/excel/excel.tablecollection)|
| `onFiltered` | <span data-ttu-id="bf436-146">Se produit lorsque le filtre est appliqué sur un objet.</span><span class="sxs-lookup"><span data-stu-id="bf436-146">Occurs when filter is applied on an object.</span></span> | <span data-ttu-id="bf436-147">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="bf436-147">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onFormatChanged` | <span data-ttu-id="bf436-148">Se produit lorsque le format est modifié sur une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="bf436-148">Occurs when format changed on a specific worksheet.</span></span> | <span data-ttu-id="bf436-149">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="bf436-149">[**WorksheetCollection**](/javascript/api/excel/excel.worksheet), [**Worksheet**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="bf436-150">Se produit lorsque la sélection change sur n’importe quelle feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="bf436-150">Occurs when the selection changes on any worksheet.</span></span> | [**<span data-ttu-id="bf436-151">WorksheetCollection</span><span class="sxs-lookup"><span data-stu-id="bf436-151">worksheetCollection</span></span>**](/javascript/api/excel/excel.worksheetcollection) |

### <a name="event-triggers"></a><span data-ttu-id="bf436-152">Déclencheurs d’événements</span><span class="sxs-lookup"><span data-stu-id="bf436-152">Event triggers</span></span>

<span data-ttu-id="bf436-153">Événements au sein d’un classeur Excel pouvant être déclenchés par :</span><span class="sxs-lookup"><span data-stu-id="bf436-153">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="bf436-154">Interaction de l’utilisateur via l’interface utilisateur Excel (IU) modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="bf436-154">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="bf436-155">Complément (JavaScript) Office modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="bf436-155">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="bf436-156">Complément VBA (macro) modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="bf436-156">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="bf436-157">Toute modification conforme aux comportements par défaut d’Excel déclenche les événements correspondants dans un classeur.</span><span class="sxs-lookup"><span data-stu-id="bf436-157">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="bf436-158">Cycle de vie d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="bf436-158">Lifecycle of an event handler</span></span>

<span data-ttu-id="bf436-159">Un gestionnaire d’événements est créé lorsqu’un complément inscrit le gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="bf436-159">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="bf436-160">Il est détruit lorsque le complément annule l’inscription du gestionnaire d’événements ou lorsque le complément est actualisé, rechargé ou fermé.</span><span class="sxs-lookup"><span data-stu-id="bf436-160">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="bf436-161">Les gestionnaires d’événements ne sont pas conservés dans le fichier Excel ou entre des sessions avec Excel Online.</span><span class="sxs-lookup"><span data-stu-id="bf436-161">Event handlers do not persist as part of the Excel file, or across sessions with Excel Online.</span></span>

> [!CAUTION]
> <span data-ttu-id="bf436-162">Lorsqu’un objet dans lequel des événements sont inscrits est supprimé (par exemple, un tableau avec un événement `onChanged`), le gestionnaire d’événements n’est plus déclenché mais reste en mémoire jusqu’à ce que le complément ou la session Excel soit actualisé(e) ou se ferme.</span><span class="sxs-lookup"><span data-stu-id="bf436-162">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="bf436-163">Événements et co-création</span><span class="sxs-lookup"><span data-stu-id="bf436-163">Events and coauthoring</span></span>

<span data-ttu-id="bf436-p104">Avec la [co-création](co-authoring-in-excel-add-ins.md), plusieurs personnes peuvent travailler ensemble et modifier le même classeur Excel simultanément. Pour les événements pouvant être déclenchés par un co-auteur, tels que `onChanged`, l’objet **Event** correspondant contient une propriété **source** qui indique si l’événement a été déclenché localement par l’utilisateur actuel (`event.source = Local`) ou par le co-auteur à distance (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="bf436-p104">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="bf436-166">Inscription d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="bf436-166">Register an event handler</span></span>

<span data-ttu-id="bf436-p105">L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onChanged` dans la feuille de calcul **Sample**. Le code indique que la fonction `handleDataChange` doit être exécutée lorsque les données de la feuille de calcul sont modifiées.</span><span class="sxs-lookup"><span data-stu-id="bf436-p105">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="bf436-169">Gestion d’un événement</span><span class="sxs-lookup"><span data-stu-id="bf436-169">Handle an event</span></span>

<span data-ttu-id="bf436-p106">Comme indiqué dans l’exemple précédent, lorsque vous inscrivez un gestionnaire d’événements, vous indiquez la fonction devant être exécutée lorsque l’événement spécifié se produit. Vous pouvez créer cette fonction pour effectuer n’importe quelle action nécessaire à votre scénario. L’exemple de code suivant montre une fonction de gestionnaire d’événements qui écrit simplement des informations sur l’événement dans la console.</span><span class="sxs-lookup"><span data-stu-id="bf436-p106">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span>

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="bf436-173">Suppression d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="bf436-173">Remove an event handler</span></span>

<span data-ttu-id="bf436-p107">L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onSelectionChanged` dans la feuille de calcul **Sample** et définit la fonction `handleSelectionChange` qui est exécutée lorsqu’un événement se produit. Il définit également la fonction `remove()` pouvant être appelée par la suite pour supprimer ce gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="bf436-p107">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs. It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="bf436-176">Activation et désactivation d’événements</span><span class="sxs-lookup"><span data-stu-id="bf436-176">Enable and disable events</span></span>

<span data-ttu-id="bf436-177">La performance d’un complément peut être améliorée en désactivant les événements.</span><span class="sxs-lookup"><span data-stu-id="bf436-177">The performance of an add-in may be improved by disabling events.</span></span>
<span data-ttu-id="bf436-178">Par exemple, il se peut que votre application ne doive jamais recevoir d’événements, ou elle peut ignorer des événements lors de modifications par lots de plusieurs entités.</span><span class="sxs-lookup"><span data-stu-id="bf436-178">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="bf436-179">Les événements sont activés et désactivés au niveau [runtime](/javascript/api/excel/excel.runtime).</span><span class="sxs-lookup"><span data-stu-id="bf436-179">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="bf436-180">La propriété `enableEvents` détermine si les événements sont déclenchés et leurs gestionnaires activés.</span><span class="sxs-lookup"><span data-stu-id="bf436-180">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="bf436-181">L’exemple de code suivant montre comment activer et désactiver des événements.</span><span class="sxs-lookup"><span data-stu-id="bf436-181">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="bf436-182">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="bf436-182">See also</span></span>

- [<span data-ttu-id="bf436-183">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="bf436-183">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
