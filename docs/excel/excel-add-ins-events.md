---
title: Utilisation d’événements à l’aide de l’API JavaScript pour Excel
description: Liste des événements pour Excel objets JavaScript. Cela inclut des informations sur l’utilisation des handlers d’événements et les modèles associés.
ms.date: 06/02/2021
localization_priority: Normal
ms.openlocfilehash: 50d9c4c2a15a955f0a96c70464fa0165625ea6f8
ms.sourcegitcommit: 17b5a076375bc5dc3f91d3602daeb7535d67745d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/06/2021
ms.locfileid: "52783504"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="eacf9-104">Utilisation d’événements à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="eacf9-104">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="eacf9-105">Cet article décrit des concepts importants relatifs à l’utilisation des événements dans Excel et fournit des exemples de code montrant comment inscrire des gestionnaires d’événements, gérer des événements et supprimer des gestionnaires d’événements à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="eacf9-105">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span>

## <a name="events-in-excel"></a><span data-ttu-id="eacf9-106">Événements dans Excel</span><span class="sxs-lookup"><span data-stu-id="eacf9-106">Events in Excel</span></span>

<span data-ttu-id="eacf9-p102">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span><span class="sxs-lookup"><span data-stu-id="eacf9-p102">Each time certain types of changes occur in an Excel workbook, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.</span></span>

| <span data-ttu-id="eacf9-110">Événement</span><span class="sxs-lookup"><span data-stu-id="eacf9-110">Event</span></span> | <span data-ttu-id="eacf9-111">Description</span><span class="sxs-lookup"><span data-stu-id="eacf9-111">Description</span></span> | <span data-ttu-id="eacf9-112">Objets pris en charge</span><span class="sxs-lookup"><span data-stu-id="eacf9-112">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onActivated` | <span data-ttu-id="eacf9-113">Se produit lorsqu’un objet est activé.</span><span class="sxs-lookup"><span data-stu-id="eacf9-113">Occurs when an object is activated.</span></span> | <span data-ttu-id="eacf9-114">[**Chart**](/javascript/api/excel/excel.chart#onactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onactivated), [**Shape**](/javascript/api/excel/excel.shape#onactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#onactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onactivated)</span><span class="sxs-lookup"><span data-stu-id="eacf9-114">[**Chart**](/javascript/api/excel/excel.chart#onactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#onactivated), [**Shape**](/javascript/api/excel/excel.shape#onactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#onactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onactivated)</span></span> |
| `onAdded` | <span data-ttu-id="eacf9-115">Se produit lorsqu’un objet est ajouté à la collection.</span><span class="sxs-lookup"><span data-stu-id="eacf9-115">Occurs when an object is added to the collection.</span></span> | <span data-ttu-id="eacf9-116">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#onadded), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#onadded), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onadded), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onadded)</span><span class="sxs-lookup"><span data-stu-id="eacf9-116">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#onadded), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#onadded), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onadded), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onadded)</span></span> |
| `onAutoSaveSettingChanged` | <span data-ttu-id="eacf9-117">Se produit lorsque le paramètre de `autoSave` est modifié dans le classeur.</span><span class="sxs-lookup"><span data-stu-id="eacf9-117">Occurs when the `autoSave` setting is changed on the workbook.</span></span> | [<span data-ttu-id="eacf9-118">**Classeur**</span><span class="sxs-lookup"><span data-stu-id="eacf9-118">**Workbook**</span></span>](/javascript/api/excel/excel.workbook#onautosavesettingchanged) |
| `onCalculated` | <span data-ttu-id="eacf9-119">Se produit lorsqu’une feuille de calcul a terminé un calcul (ou que toutes les feuilles de calcul de la collection ont terminé).</span><span class="sxs-lookup"><span data-stu-id="eacf9-119">Occurs when a worksheet has finished calculation (or all the worksheets of the collection have finished).</span></span> | <span data-ttu-id="eacf9-120">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncalculated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncalculated)</span><span class="sxs-lookup"><span data-stu-id="eacf9-120">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncalculated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncalculated)</span></span> |
| `onChanged` | <span data-ttu-id="eacf9-121">Se produit lorsque les données de cellules individuelles ou de commentaires ont changé.</span><span class="sxs-lookup"><span data-stu-id="eacf9-121">Occurs when the data of individual cells or comments has changed.</span></span> | <span data-ttu-id="eacf9-122">[**CommentCollection**](/javascript/api/excel/excel.commentcollection#onchanged), [**Table**](/javascript/api/excel/excel.table#onchanged), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onchanged)</span><span class="sxs-lookup"><span data-stu-id="eacf9-122">[**CommentCollection**](/javascript/api/excel/excel.commentcollection#onchanged), [**Table**](/javascript/api/excel/excel.table#onchanged), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onchanged)</span></span> |
| `onColumnSorted` | <span data-ttu-id="eacf9-123">Se produit lorsqu’une ou plusieurs colonnes ont été triées.</span><span class="sxs-lookup"><span data-stu-id="eacf9-123">Occurs when one or more columns have been sorted.</span></span> <span data-ttu-id="eacf9-124">Ce problème se produit en raison de l’opération de tri de gauche à droite.</span><span class="sxs-lookup"><span data-stu-id="eacf9-124">This happens as the result of a left-to-right sort operation.</span></span> | <span data-ttu-id="eacf9-125">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncolumnsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)</span><span class="sxs-lookup"><span data-stu-id="eacf9-125">[**Worksheet**](/javascript/api/excel/excel.worksheet#oncolumnsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)</span></span> |
| `onDataChanged` | <span data-ttu-id="eacf9-126">Se produit lors de la modification des données ou de la mise en forme dans la liaison.</span><span class="sxs-lookup"><span data-stu-id="eacf9-126">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="eacf9-127">**Binding**</span><span class="sxs-lookup"><span data-stu-id="eacf9-127">**Binding**</span></span>](/javascript/api/excel/excel.binding#ondatachanged) |
| `onDeactivated` | <span data-ttu-id="eacf9-128">Se produit lorsqu’un objet est désactivé.</span><span class="sxs-lookup"><span data-stu-id="eacf9-128">Occurs when an object is deactivated.</span></span> | <span data-ttu-id="eacf9-129">[**Chart**](/javascript/api/excel/excel.chart#ondeactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeactivated), [**Shape**](/javascript/api/excel/excel.shape#ondeactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#ondeactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeactivated)</span><span class="sxs-lookup"><span data-stu-id="eacf9-129">[**Chart**](/javascript/api/excel/excel.chart#ondeactivated), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeactivated), [**Shape**](/javascript/api/excel/excel.shape#ondeactivated), [**Worksheet**](/javascript/api/excel/excel.worksheet#ondeactivated), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeactivated)</span></span> |
| `onDeleted` | <span data-ttu-id="eacf9-130">Se produit lorsqu’un objet est supprimé de la collection.</span><span class="sxs-lookup"><span data-stu-id="eacf9-130">Occurs when an object is deleted from the collection.</span></span> | <span data-ttu-id="eacf9-131">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeleted), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#ondeleted), [**TableCollection**](/javascript/api/excel/excel.tablecollection#ondeleted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeleted)</span><span class="sxs-lookup"><span data-stu-id="eacf9-131">[**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeleted), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#ondeleted), [**TableCollection**](/javascript/api/excel/excel.tablecollection#ondeleted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeleted)</span></span> |
| `onFormatChanged` | <span data-ttu-id="eacf9-132">Se produit lorsque le format est modifié sur une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="eacf9-132">Occurs when the format is changed on a worksheet.</span></span> | <span data-ttu-id="eacf9-133">[**Worksheet**](/javascript/api/excel/excel.worksheet#onformatchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onformatchanged)</span><span class="sxs-lookup"><span data-stu-id="eacf9-133">[**Worksheet**](/javascript/api/excel/excel.worksheet#onformatchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onformatchanged)</span></span> |
| `onRowSorted` | <span data-ttu-id="eacf9-134">Se produit lorsqu’une ou plusieurs lignes ont été triées.</span><span class="sxs-lookup"><span data-stu-id="eacf9-134">Occurs when one or more rows have been sorted.</span></span> <span data-ttu-id="eacf9-135">Cela se produit en raison d’une opération de tri de haut en bas.</span><span class="sxs-lookup"><span data-stu-id="eacf9-135">This happens as the result of a top-to-bottom sort operation.</span></span> | <span data-ttu-id="eacf9-136">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowsorted)</span><span class="sxs-lookup"><span data-stu-id="eacf9-136">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowsorted), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowsorted)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="eacf9-137">Se produit lorsque la cellule active ou la plage sélectionnée est modifiée.</span><span class="sxs-lookup"><span data-stu-id="eacf9-137">Occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="eacf9-138">[**Binding**](/javascript/api/excel/excel.binding#onselectionchanged), [**Table**](/javascript/api/excel/excel.table#onselectionchanged), [**Workbook**](/javascript/api/excel/excel.workbook#onselectionchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onselectionchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)</span><span class="sxs-lookup"><span data-stu-id="eacf9-138">[**Binding**](/javascript/api/excel/excel.binding#onselectionchanged), [**Table**](/javascript/api/excel/excel.table#onselectionchanged), [**Workbook**](/javascript/api/excel/excel.workbook#onselectionchanged), [**Worksheet**](/javascript/api/excel/excel.worksheet#onselectionchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)</span></span> |
| `onRowHiddenChanged` | <span data-ttu-id="eacf9-139">Se produit lorsque l’état de ligne masquée change sur une feuille de calcul spécifique.</span><span class="sxs-lookup"><span data-stu-id="eacf9-139">Occurs when the row-hidden state changes on a specific worksheet.</span></span> | <span data-ttu-id="eacf9-140">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowhiddenchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)</span><span class="sxs-lookup"><span data-stu-id="eacf9-140">[**Worksheet**](/javascript/api/excel/excel.worksheet#onrowhiddenchanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)</span></span> |
| `onSettingsChanged` | <span data-ttu-id="eacf9-141">Se produit lorsque les paramètres dans le document sont modifiés.</span><span class="sxs-lookup"><span data-stu-id="eacf9-141">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="eacf9-142">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="eacf9-142">**SettingCollection**</span></span>](/javascript/api/excel/excel.settingcollection#onsettingschanged) |
| `onSingleClicked` | <span data-ttu-id="eacf9-143">Se produit lorsque l’opération clic gauche/tape se produit dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="eacf9-143">Occurs when left-clicked/tapped action occurs in the worksheet.</span></span> | <span data-ttu-id="eacf9-144">[**Worksheet**](/javascript/api/excel/excel.worksheet#onsingleclicked), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)</span><span class="sxs-lookup"><span data-stu-id="eacf9-144">[**Worksheet**](/javascript/api/excel/excel.worksheet#onsingleclicked), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)</span></span> |

### <a name="events-in-preview"></a><span data-ttu-id="eacf9-145">Événements en préversion</span><span class="sxs-lookup"><span data-stu-id="eacf9-145">Events in preview</span></span>

> [!NOTE]
> <span data-ttu-id="eacf9-146">Les événements suivants sont actuellement disponibles uniquement en préversion publique.</span><span class="sxs-lookup"><span data-stu-id="eacf9-146">The following events are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| <span data-ttu-id="eacf9-147">Événement</span><span class="sxs-lookup"><span data-stu-id="eacf9-147">Event</span></span> | <span data-ttu-id="eacf9-148">Description</span><span class="sxs-lookup"><span data-stu-id="eacf9-148">Description</span></span> | <span data-ttu-id="eacf9-149">Objets pris en charge</span><span class="sxs-lookup"><span data-stu-id="eacf9-149">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onFiltered` | <span data-ttu-id="eacf9-150">Se produit lorsqu’un filtre est appliqué à un objet.</span><span class="sxs-lookup"><span data-stu-id="eacf9-150">Occurs when a filter is applied to an object.</span></span> | <span data-ttu-id="eacf9-151">[**Table**](/javascript/api/excel/excel.table#onfiltered), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onfiltered), [**Worksheet**](/javascript/api/excel/excel.worksheet#onfiltered), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onfiltered)</span><span class="sxs-lookup"><span data-stu-id="eacf9-151">[**Table**](/javascript/api/excel/excel.table#onfiltered), [**TableCollection**](/javascript/api/excel/excel.tablecollection#onfiltered), [**Worksheet**](/javascript/api/excel/excel.worksheet#onfiltered), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onfiltered)</span></span> |
| `onFormulaChanged` | <span data-ttu-id="eacf9-152">Se produit lorsqu’une formule est modifiée.</span><span class="sxs-lookup"><span data-stu-id="eacf9-152">Occurs when a formula is changed.</span></span> | <span data-ttu-id="eacf9-153">[**Worksheet**](/javascript/api/excel/excel.worksheet#onFormulaChanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onFormulaChanged)</span><span class="sxs-lookup"><span data-stu-id="eacf9-153">[**Worksheet**](/javascript/api/excel/excel.worksheet#onFormulaChanged), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onFormulaChanged)</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="eacf9-154">Déclencheurs d’événements</span><span class="sxs-lookup"><span data-stu-id="eacf9-154">Event triggers</span></span>

<span data-ttu-id="eacf9-155">Événements au sein d’un classeur Excel pouvant être déclenchés par :</span><span class="sxs-lookup"><span data-stu-id="eacf9-155">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="eacf9-156">Interaction de l’utilisateur via l’interface utilisateur Excel (IU) modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="eacf9-156">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="eacf9-157">Complément (JavaScript) Office modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="eacf9-157">Office Add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="eacf9-158">Complément VBA (macro) modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="eacf9-158">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="eacf9-159">Toute modification conforme aux comportements par défaut d’Excel déclenche les événements correspondants dans un classeur.</span><span class="sxs-lookup"><span data-stu-id="eacf9-159">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="eacf9-160">Cycle de vie d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="eacf9-160">Lifecycle of an event handler</span></span>

<span data-ttu-id="eacf9-161">Un gestionnaire d’événements est créé lorsqu’un complément inscrit le gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="eacf9-161">An event handler is created when an add-in registers the event handler.</span></span> <span data-ttu-id="eacf9-162">Il est détruit lorsque le complément annule l’inscription du gestionnaire d’événements ou lorsque le complément est actualisé, rechargé ou fermé.</span><span class="sxs-lookup"><span data-stu-id="eacf9-162">It is destroyed when the add-in unregisters the event handler or when the add-in is refreshed, reloaded, or closed.</span></span> <span data-ttu-id="eacf9-163">Les gestionnaires d’événements ne sont pas conservés dans le fichier Excel ou entre des sessions avec Excel sur le web.</span><span class="sxs-lookup"><span data-stu-id="eacf9-163">Event handlers do not persist as part of the Excel file, or across sessions with Excel on the web.</span></span>

> [!CAUTION]
> <span data-ttu-id="eacf9-164">Lorsqu’un objet dans lequel des événements sont inscrits est supprimé (par exemple, un tableau avec un événement `onChanged`), le gestionnaire d’événements n’est plus déclenché mais reste en mémoire jusqu’à ce que le complément ou la session Excel soit actualisé(e) ou se ferme.</span><span class="sxs-lookup"><span data-stu-id="eacf9-164">When an object to which events are registered is deleted (e.g., a table with an `onChanged` event registered), the event handler no longer triggers but remains in memory until the add-in or Excel session refreshes or closes.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="eacf9-165">Événements et co-création</span><span class="sxs-lookup"><span data-stu-id="eacf9-165">Events and coauthoring</span></span>

<span data-ttu-id="eacf9-p107">Avec la [co-création](co-authoring-in-excel-add-ins.md), plusieurs personnes peuvent travailler ensemble et modifier le même classeur Excel simultanément. Pour les événements pouvant être déclenchés par un co-auteur, tels que `onChanged`, l’objet **Event** correspondant contient une propriété **source** qui indique si l’événement a été déclenché localement par l’utilisateur actuel (`event.source = Local`) ou par le co-auteur à distance (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="eacf9-p107">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="eacf9-168">Inscription d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="eacf9-168">Register an event handler</span></span>

<span data-ttu-id="eacf9-p108">L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onChanged` dans la feuille de calcul **Sample**. Le code indique que la fonction `handleChange` doit être exécutée lorsque les données de la feuille de calcul sont modifiées.</span><span class="sxs-lookup"><span data-stu-id="eacf9-p108">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**. The code specifies that when data changes in that worksheet, the `handleChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="eacf9-171">Gestion d’un événement</span><span class="sxs-lookup"><span data-stu-id="eacf9-171">Handle an event</span></span>

<span data-ttu-id="eacf9-p109">Comme indiqué dans l’exemple précédent, lorsque vous inscrivez un gestionnaire d’événements, vous indiquez la fonction devant être exécutée lorsque l’événement spécifié se produit. Vous pouvez créer cette fonction pour effectuer n’importe quelle action nécessaire à votre scénario. L’exemple de code suivant montre une fonction de gestionnaire d’événements qui écrit simplement des informations sur l’événement dans la console.</span><span class="sxs-lookup"><span data-stu-id="eacf9-p109">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs. You can design that function to perform whatever actions your scenario requires. The following code sample shows an event handler function that simply writes information about the event to the console.</span></span>

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="eacf9-175">Suppression d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="eacf9-175">Remove an event handler</span></span>

<span data-ttu-id="eacf9-176">L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onSelectionChanged` dans la feuille de calcul **Sample** et définit la fonction `handleSelectionChange` qui est exécutée lorsqu’un événement se produit.</span><span class="sxs-lookup"><span data-stu-id="eacf9-176">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="eacf9-177">Il définit également la fonction `remove()` pouvant être appelée par la suite pour supprimer ce gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="eacf9-177">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span> <span data-ttu-id="eacf9-178">Notez que `RequestContext` l’utilisation pour créer le handler d’événements est nécessaire pour le supprimer.</span><span class="sxs-lookup"><span data-stu-id="eacf9-178">Note that the `RequestContext` used to create the event handler is needed to remove it.</span></span> 

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="eacf9-179">Activation et désactivation d’événements</span><span class="sxs-lookup"><span data-stu-id="eacf9-179">Enable and disable events</span></span>

<span data-ttu-id="eacf9-180">La performance d’un complément peut être améliorée en désactivant les événements.</span><span class="sxs-lookup"><span data-stu-id="eacf9-180">The performance of an add-in may be improved by disabling events.</span></span>
<span data-ttu-id="eacf9-181">Par exemple, il se peut que votre application ne doive jamais recevoir d’événements, ou elle peut ignorer des événements lors de modifications par lots de plusieurs entités.</span><span class="sxs-lookup"><span data-stu-id="eacf9-181">For example, your app might never need to receive events, or it could ignore events while performing batch-edits of multiple entities.</span></span>

<span data-ttu-id="eacf9-182">Les événements sont activés et désactivés au niveau [runtime](/javascript/api/excel/excel.runtime).</span><span class="sxs-lookup"><span data-stu-id="eacf9-182">Events are enabled and disabled at the [runtime](/javascript/api/excel/excel.runtime) level.</span></span>
<span data-ttu-id="eacf9-183">La propriété `enableEvents` détermine si les événements sont déclenchés et leurs gestionnaires activés.</span><span class="sxs-lookup"><span data-stu-id="eacf9-183">The `enableEvents` property determines if events are fired and their handlers are activated.</span></span>

<span data-ttu-id="eacf9-184">L’exemple de code suivant montre comment activer et désactiver des événements.</span><span class="sxs-lookup"><span data-stu-id="eacf9-184">The following code sample shows how to toggle events on and off.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="eacf9-185">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="eacf9-185">See also</span></span>

- [<span data-ttu-id="eacf9-186">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="eacf9-186">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
