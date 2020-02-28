---
title: Utilisation d’événements à l’aide de l’API JavaScript pour Excel
description: Liste d’événements pour les objets JavaScript Excel. Cela inclut des informations sur l’utilisation des gestionnaires d’événements et les modèles associés.
ms.date: 02/11/2020
localization_priority: Normal
ms.openlocfilehash: f1a1faf9acc370e7183a078aeeba34019e54900f
ms.sourcegitcommit: d85efbf41a3382ca7d3ab08f2c3f0664d4b26c53
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/28/2020
ms.locfileid: "42327774"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="4c772-104">Utilisation d’événements à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="4c772-104">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="4c772-105">Cet article décrit des concepts importants relatifs à l’utilisation des événements dans Excel et fournit des exemples de code montrant comment inscrire des gestionnaires d’événements, gérer des événements et supprimer des gestionnaires d’événements à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="4c772-105">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span>

## <a name="events-in-excel"></a><span data-ttu-id="4c772-106">Événements dans Excel</span><span class="sxs-lookup"><span data-stu-id="4c772-106">Events in Excel</span></span>

<span data-ttu-id="4c772-p102">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span><span class="sxs-lookup"><span data-stu-id="4c772-p102">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="4c772-110">Événement</span><span class="sxs-lookup"><span data-stu-id="4c772-110">Event</span></span> | <span data-ttu-id="4c772-111">Description</span><span class="sxs-lookup"><span data-stu-id="4c772-111">Description</span></span> | <span data-ttu-id="4c772-112">Objets pris en charge</span><span class="sxs-lookup"><span data-stu-id="4c772-112">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="4c772-113">Se produit lorsqu’un objet est activé.</span><span class="sxs-lookup"><span data-stu-id="4c772-113">Occurs when an object is activated.</span></span> | <span data-ttu-id="4c772-114">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Shape**](/javascript/api/excel/excel.shape), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="4c772-114">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Shape**](/javascript/api/excel/excel.shape), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onAdded` | <span data-ttu-id="4c772-115">Se produit lorsqu’un objet est ajouté à la collection.</span><span class="sxs-lookup"><span data-stu-id="4c772-115">Occurs when an object is added to the collection.</span></span> | <span data-ttu-id="4c772-116">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="4c772-116">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onAutoSaveSettingChanged` | <span data-ttu-id="4c772-117">Se produit lorsque le paramètre de `autoSave` est modifié dans le classeur.</span><span class="sxs-lookup"><span data-stu-id="4c772-117">Occurs when the `autoSave` setting is changed on the workbook.</span></span> | [<span data-ttu-id="4c772-118">**Classeur**</span><span class="sxs-lookup"><span data-stu-id="4c772-118">**Workbook**</span></span>](/javascript/api/excel/excel.workbook) |
| `onCalculated` | <span data-ttu-id="4c772-119">Se produit lorsqu’une feuille de calcul a terminé un calcul (ou que toutes les feuilles de calcul de la collection ont terminé).</span><span class="sxs-lookup"><span data-stu-id="4c772-119">Occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="4c772-120">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="4c772-120">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="4c772-121">Se produit lorsque les données au sein des cellules sont modifiées.</span><span class="sxs-lookup"><span data-stu-id="4c772-121">Occurs when data within cells is changed.</span></span> | <span data-ttu-id="4c772-122">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="4c772-122">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onColumnSorted` | <span data-ttu-id="4c772-123">Se produit lorsqu’une ou plusieurs colonnes ont été triées.</span><span class="sxs-lookup"><span data-stu-id="4c772-123">Occurs when one or more columns have been sorted.</span></span> <span data-ttu-id="4c772-124">Ce problème se produit en raison de l’opération de tri de gauche à droite.</span><span class="sxs-lookup"><span data-stu-id="4c772-124">This happens as the result of a left-to-right sort operation.</span></span> | <span data-ttu-id="4c772-125">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="4c772-125">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="4c772-126">Se produit lors de la modification des données ou de la mise en forme dans la liaison.</span><span class="sxs-lookup"><span data-stu-id="4c772-126">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="4c772-127">**Binding**</span><span class="sxs-lookup"><span data-stu-id="4c772-127">**Binding**</span></span>](/javascript/api/excel/excel.binding) |
| `onDeactivated` | <span data-ttu-id="4c772-128">Se produit lorsqu’un objet est désactivé.</span><span class="sxs-lookup"><span data-stu-id="4c772-128">Occurs when an object is deactivated.</span></span> | <span data-ttu-id="4c772-129">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Shape**](/javascript/api/excel/excel.shape), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="4c772-129">[**Chart**](/javascript/api/excel/excel.chart), [**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**Shape**](/javascript/api/excel/excel.shape), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onDeleted` | <span data-ttu-id="4c772-130">Se produit lorsqu’un objet est supprimé de la collection.</span><span class="sxs-lookup"><span data-stu-id="4c772-130">Occurs when an object is deleted from the collection.</span></span> | <span data-ttu-id="4c772-131">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="4c772-131">[**ChartCollection**](/javascript/api/excel/excel.chartcollection), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onFormatChanged` | <span data-ttu-id="4c772-132">Se produit lorsque le format est modifié sur une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="4c772-132">Occurs when the format is changed on a worksheet.</span></span> | <span data-ttu-id="4c772-133">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="4c772-133">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onRowSorted` | <span data-ttu-id="4c772-134">Se produit lorsqu’une ou plusieurs lignes ont été triées.</span><span class="sxs-lookup"><span data-stu-id="4c772-134">Occurs when one or more rows have been sorted.</span></span> <span data-ttu-id="4c772-135">Cela se produit en raison d’une opération de tri de haut en bas.</span><span class="sxs-lookup"><span data-stu-id="4c772-135">This happens as the result of a top-to-bottom sort operation.</span></span> | <span data-ttu-id="4c772-136">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="4c772-136">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="4c772-137">Se produit lorsque la cellule active ou la plage sélectionnée est modifiée.</span><span class="sxs-lookup"><span data-stu-id="4c772-137">Occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="4c772-138">[**Binding**](/javascript/api/excel/excel.binding), [**Table**](/javascript/api/excel/excel.table), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="4c772-138">[**Binding**](/javascript/api/excel/excel.binding), [**Table**](/javascript/api/excel/excel.table),  [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="4c772-139">Se produit lorsque les paramètres dans le document sont modifiés.</span><span class="sxs-lookup"><span data-stu-id="4c772-139">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="4c772-140">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="4c772-140">**SettingCollection**</span></span>](/javascript/api/excel/excel.settingcollection) |
| `onSingleClicked` | <span data-ttu-id="4c772-141">Se produit lorsque l’opération clic gauche/tape se produit dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="4c772-141">Occurs when left-clicked/tapped action occurs in the worksheet.</span></span> | <span data-ttu-id="4c772-142">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="4c772-142">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |

> [!WARNING]
> <span data-ttu-id="4c772-143">`onSelectionChanged` est actuellement instable.</span><span class="sxs-lookup"><span data-stu-id="4c772-143">`onSelectionChanged` is currently unstable.</span></span> <span data-ttu-id="4c772-144">Il existe une solution de contournement pour utiliser `onSelectionChanged` de façon fiable.</span><span class="sxs-lookup"><span data-stu-id="4c772-144">There is a workaround to reliably use `onSelectionChanged`.</span></span> <span data-ttu-id="4c772-145">Ajoutez le code suivant dans la section `<head>` de votre page d’accueil HTML :</span><span class="sxs-lookup"><span data-stu-id="4c772-145">Add the following code to the `<head>` section of your HTML home page:</span></span>
>
> ```HTML
> <script> MutationObserver=null; </script>
> ```
>
> <span data-ttu-id="4c772-146">Une discussion complète sur ce problème est disponible sur le [référentiel GitHub office-js](https://github.com/OfficeDev/office-js/issues/533).</span><span class="sxs-lookup"><span data-stu-id="4c772-146">A full discussion of the issue can be found on the [office-js GitHub repo](https://github.com/OfficeDev/office-js/issues/533).</span></span>

### <a name="events-in-preview"></a><span data-ttu-id="4c772-147">Événements en préversion</span><span class="sxs-lookup"><span data-stu-id="4c772-147">Events in preview</span></span>

> [!NOTE]
> <span data-ttu-id="4c772-148">Les événements suivants sont actuellement disponibles uniquement en préversion publique.</span><span class="sxs-lookup"><span data-stu-id="4c772-148">The following events are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| <span data-ttu-id="4c772-149">Événement</span><span class="sxs-lookup"><span data-stu-id="4c772-149">Event</span></span> | <span data-ttu-id="4c772-150">Description</span><span class="sxs-lookup"><span data-stu-id="4c772-150">Description</span></span> | <span data-ttu-id="4c772-151">Objets pris en charge</span><span class="sxs-lookup"><span data-stu-id="4c772-151">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onFiltered` | <span data-ttu-id="4c772-152">Se produit lorsqu’un filtre est appliqué à un objet.</span><span class="sxs-lookup"><span data-stu-id="4c772-152">Occurs when a filter is applied to an object.</span></span> | <span data-ttu-id="4c772-153">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="4c772-153">[**Table**](/javascript/api/excel/excel.table), [**TableCollection**](/javascript/api/excel/excel.tablecollection), [**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |
| `onRowHiddenChanged` | <span data-ttu-id="4c772-154">Se produit lorsque l’état de ligne masquée change sur une feuille de calcul spécifique.</span><span class="sxs-lookup"><span data-stu-id="4c772-154">Occurs when the row-hidden state changes on a specific worksheet.</span></span> | <span data-ttu-id="4c772-155">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span><span class="sxs-lookup"><span data-stu-id="4c772-155">[**Worksheet**](/javascript/api/excel/excel.worksheet), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="4c772-156">Déclencheurs d’événements</span><span class="sxs-lookup"><span data-stu-id="4c772-156">Event triggers</span></span>

<span data-ttu-id="4c772-157">Événements au sein d’un classeur Excel pouvant être déclenchés par :</span><span class="sxs-lookup"><span data-stu-id="4c772-157">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="4c772-158">Interaction de l’utilisateur via l’interface utilisateur Excel (IU) modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="4c772-158">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="4c772-159">Complément (JavaScript) Office modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="4c772-159">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="4c772-160">Complément VBA (macro) modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="4c772-160">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="4c772-161">Toute modification conforme aux comportements par défaut d’Excel déclenche les événements correspondants dans un classeur.</span><span class="sxs-lookup"><span data-stu-id="4c772-161">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="4c772-162">Cycle de vie d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="4c772-162">Lifecycle of an event handler</span></span>

<span data-ttu-id="4c772-163">Un gestionnaire d’événements est créé lorsqu’un complément inscrit le gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="4c772-163">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="4c772-164">Il est détruit lorsque le complément annule l’inscription du gestionnaire d’événements ou lorsque le complément est actualisé, rechargé ou fermé.</span><span class="sxs-lookup"><span data-stu-id="4c772-164">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="4c772-165">Les gestionnaires d’événements ne sont pas conservés dans le fichier Excel ou entre des sessions avec Excel sur le web.</span><span class="sxs-lookup"><span data-stu-id="4c772-165">Event handlers do not persist as part of the Excel file, or across sessions with Excel on the web.</span></span>

> [!CAUTION]
> <span data-ttu-id="4c772-166">Lorsqu’un objet dans lequel des événements sont inscrits est supprimé (par exemple, un tableau avec un événement `onChanged`), le gestionnaire d’événements n’est plus déclenché mais reste en mémoire jusqu’à ce que le complément ou la session Excel soit actualisé(e) ou se ferme.</span><span class="sxs-lookup"><span data-stu-id="4c772-166">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="4c772-167">Événements et co-création</span><span class="sxs-lookup"><span data-stu-id="4c772-167">Events and coauthoring</span></span>

<span data-ttu-id="4c772-p108">Avec la [co-création](co-authoring-in-excel-add-ins.md), plusieurs personnes peuvent travailler ensemble et modifier le même classeur Excel simultanément. Pour les événements pouvant être déclenchés par un co-auteur, tels que `onChanged`, l’objet **Event** correspondant contient une propriété **source** qui indique si l’événement a été déclenché localement par l’utilisateur actuel (`event.source = Local`) ou par le co-auteur à distance (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="4c772-p108">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="4c772-170">Inscription d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="4c772-170">Register an event handler</span></span>

<span data-ttu-id="4c772-p109">L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onChanged` dans la feuille de calcul **Sample**. Le code indique que la fonction `handleDataChange` doit être exécutée lorsque les données de la feuille de calcul sont modifiées.</span><span class="sxs-lookup"><span data-stu-id="4c772-p109">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="4c772-173">Gestion d’un événement</span><span class="sxs-lookup"><span data-stu-id="4c772-173">Handle an event</span></span>

<span data-ttu-id="4c772-p110">Comme indiqué dans l’exemple précédent, lorsque vous inscrivez un gestionnaire d’événements, vous indiquez la fonction devant être exécutée lorsque l’événement spécifié se produit. Vous pouvez créer cette fonction pour effectuer n’importe quelle action nécessaire à votre scénario. L’exemple de code suivant montre une fonction de gestionnaire d’événements qui écrit simplement des informations sur l’événement dans la console.</span><span class="sxs-lookup"><span data-stu-id="4c772-p110">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span>

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="4c772-177">Suppression d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="4c772-177">Remove an event handler</span></span>

<span data-ttu-id="4c772-178">L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onSelectionChanged` dans la feuille de calcul **Sample** et définit la fonction `handleSelectionChange` qui est exécutée lorsqu’un événement se produit.</span><span class="sxs-lookup"><span data-stu-id="4c772-178">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="4c772-179">Il définit également la fonction `remove()` pouvant être appelée par la suite pour supprimer ce gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="4c772-179">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span> <span data-ttu-id="4c772-180">Notez que le `RequestContext` utilisé pour créer le gestionnaire d’événements est nécessaire pour le supprimer.</span><span class="sxs-lookup"><span data-stu-id="4c772-180">Note that the `RequestContext` used to create the event handler is needed to remove it.</span></span> 

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="4c772-181">Activation et désactivation d’événements</span><span class="sxs-lookup"><span data-stu-id="4c772-181">Enable and disable events</span></span>

<span data-ttu-id="4c772-182">La performance d’un complément peut être améliorée en désactivant les événements.</span><span class="sxs-lookup"><span data-stu-id="4c772-182">The performance of an add-in may be improved by disabling events.</span></span>
<span data-ttu-id="4c772-183">Par exemple, il se peut que votre application ne doive jamais recevoir d’événements, ou elle peut ignorer des événements lors de modifications par lots de plusieurs entités.</span><span class="sxs-lookup"><span data-stu-id="4c772-183">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="4c772-184">Les événements sont activés et désactivés au niveau [runtime](/javascript/api/excel/excel.runtime).</span><span class="sxs-lookup"><span data-stu-id="4c772-184">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="4c772-185">La propriété `enableEvents` détermine si les événements sont déclenchés et leurs gestionnaires activés.</span><span class="sxs-lookup"><span data-stu-id="4c772-185">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="4c772-186">L’exemple de code suivant montre comment activer et désactiver des événements.</span><span class="sxs-lookup"><span data-stu-id="4c772-186">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="4c772-187">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="4c772-187">See also</span></span>

- [<span data-ttu-id="4c772-188">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="4c772-188">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
