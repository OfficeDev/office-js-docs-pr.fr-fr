---
title: Utilisation d’événements à l’aide de l’API JavaScript pour Excel
description: ''
ms.date: 05/25/2018
ms.openlocfilehash: 575e4112ed5f55356020eed8327d309fc58cd643
ms.sourcegitcommit: 9685fd83136bd2106f4c5595bda0010bc1b1950b
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/19/2018
ms.locfileid: "20596518"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="420f7-102">Utilisation d’événements à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="420f7-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="420f7-103">Cet article décrit des concepts importants relatifs à l’utilisation des événements dans Excel et fournit des exemples de code montrant comment inscrire des gestionnaires d’événements, gérer des événements et supprimer des gestionnaires d’événements à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="420f7-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

## <a name="events-in-excel"></a><span data-ttu-id="420f7-104">Événements dans Excel</span><span class="sxs-lookup"><span data-stu-id="420f7-104">Events in Excel</span></span>

<span data-ttu-id="420f7-105">Chaque fois que certains types de modifications se produisent dans un classeur Excel, une notification d’événement se déclenche.</span><span class="sxs-lookup"><span data-stu-id="420f7-105">Each time certain types of changes occur in an Excel workbook, an event notification fires.</span></span> <span data-ttu-id="420f7-106">En utilisant l’API JavaScript pour Excel, vous pouvez inscrire les gestionnaires d’événements autorisant votre complément à exécuter automatiquement une fonction désignée lorsqu’un événement spécifique se produit.</span><span class="sxs-lookup"><span data-stu-id="420f7-106">By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs.</span></span> <span data-ttu-id="420f7-107">Les événements suivants sont actuellement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="420f7-107">The following events are currently supported.</span></span>

| <span data-ttu-id="420f7-108">Événement</span><span class="sxs-lookup"><span data-stu-id="420f7-108">Event</span></span> | <span data-ttu-id="420f7-109">Description</span><span class="sxs-lookup"><span data-stu-id="420f7-109">Description</span></span> | <span data-ttu-id="420f7-110">Objets pris en charge</span><span class="sxs-lookup"><span data-stu-id="420f7-110">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="420f7-111">Événement se produisant lors de l’ajout d’un objet.</span><span class="sxs-lookup"><span data-stu-id="420f7-111">Event that occurs when an object is added.</span></span> | [<span data-ttu-id="420f7-112">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="420f7-112">**WorksheetCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onDeleted` | <span data-ttu-id="420f7-113">Événement se produisant lors de la suppression d'un objet.</span><span class="sxs-lookup"><span data-stu-id="420f7-113">Event that occurs when an object is deleted.</span></span> | [<span data-ttu-id="420f7-114">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="420f7-114">**WorksheetCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/worksheetcollection) |
| `onActivated` | <span data-ttu-id="420f7-115">Événement se produisant lors de l'activation d'un objet.</span><span class="sxs-lookup"><span data-stu-id="420f7-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="420f7-116">[**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection), [**Feuille de calcul**](https://dev.office.com/reference/add-ins/excel/worksheet)</span><span class="sxs-lookup"><span data-stu-id="420f7-116">**WorksheetCollection**, [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheetcollection)</span></span> |
| `onDeactivated` | <span data-ttu-id="420f7-117">Événement se produisant lors de la désactivation d'un objet.</span><span class="sxs-lookup"><span data-stu-id="420f7-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="420f7-118">[**WorksheetCollection**](https://dev.office.com/reference/add-ins/excel/worksheetcollection), [**Feuille de calcul**](https://dev.office.com/reference/add-ins/excel/worksheet)</span><span class="sxs-lookup"><span data-stu-id="420f7-118">**WorksheetCollection**, [Worksheet](https://dev.office.com/reference/add-ins/excel/worksheetcollection)</span></span> |
| `onChanged` | <span data-ttu-id="420f7-119">Événement se produisant lors de la modification de données dans les cellules.</span><span class="sxs-lookup"><span data-stu-id="420f7-119">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="420f7-120">[**Feuille de calcul**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Table**](https://dev.office.com/reference/add-ins/excel/table), [**TableCollection**](https://dev.office.com/reference/add-ins/excel/tablecollection)</span><span class="sxs-lookup"><span data-stu-id="420f7-120">[**Worksheet**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Table**](https://dev.office.com/reference/add-ins/excel/table), [**TableCollection**](https://dev.office.com/reference/add-ins/excel/tablecollection)</span></span> |
| `onDataChanged` | <span data-ttu-id="420f7-121">Événement se produisant lors de la modification des données ou de la mise en forme dans la liaison.</span><span class="sxs-lookup"><span data-stu-id="420f7-121">Occurs when data or formatting within the binding is changed.</span></span> | [<span data-ttu-id="420f7-122">**Liaison**</span><span class="sxs-lookup"><span data-stu-id="420f7-122">**Binding**</span></span>](https://dev.office.com/reference/add-ins/excel/binding) |
| `onSelectionChanged` | <span data-ttu-id="420f7-123">Événement se produisant lorsque la cellule active ou la plage sélectionnée est modifiée.</span><span class="sxs-lookup"><span data-stu-id="420f7-123">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="420f7-124">[**Feuille de calcul**](https://dev.office.com/reference/add-ins/excel/worksheet), [**Table**](https://dev.office.com/reference/add-ins/excel/table), [**Liaison**](https://dev.office.com/reference/add-ins/excel/binding)</span><span class="sxs-lookup"><span data-stu-id="420f7-124">**Worksheet**, [Table](https://dev.office.com/reference/add-ins/excel/worksheet), **Binding**</span></span> |
| `onSettingsChanged` | <span data-ttu-id="420f7-125">Événement qui se produit lorsque les paramètres dans le document sont modifiés.</span><span class="sxs-lookup"><span data-stu-id="420f7-125">Occurs when the Settings in the document are changed.</span></span> | [<span data-ttu-id="420f7-126">**SettingCollection**</span><span class="sxs-lookup"><span data-stu-id="420f7-126">**SettingCollection**</span></span>](https://dev.office.com/reference/add-ins/excel/settingcollection) |

### <a name="event-triggers"></a><span data-ttu-id="420f7-127">Déclencheurs d’événements</span><span class="sxs-lookup"><span data-stu-id="420f7-127">Event triggers</span></span>

<span data-ttu-id="420f7-128">Événements au sein d’un classeur Excel pouvant être déclenchés par :</span><span class="sxs-lookup"><span data-stu-id="420f7-128">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="420f7-129">Interaction de l’utilisateur via l’interface utilisateur Excel (IU) modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="420f7-129">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="420f7-130">Complément (JavaScript) Office modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="420f7-130">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="420f7-131">Complément VBA (macro) modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="420f7-131">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="420f7-132">Toute modification conforme aux comportements par défaut d’Excel déclenche les événements correspondants dans un classeur.</span><span class="sxs-lookup"><span data-stu-id="420f7-132">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="420f7-133">Cycle de vie d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="420f7-133">Lifecycle of an event handler</span></span>

<span data-ttu-id="420f7-p102">Un gestionnaire d’événements est créé lorsqu’un complément inscrit le gestionnaire d’événements et est détruit lorsque le complément désinscrit le gestionnaire d’événements ou que le complément est fermé. Les gestionnaires d’événements ne persistent pas en tant que partie du fichier Excel.</span><span class="sxs-lookup"><span data-stu-id="420f7-p102">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="420f7-136">Événements et co-création</span><span class="sxs-lookup"><span data-stu-id="420f7-136">Events and coauthoring</span></span>

<span data-ttu-id="420f7-p103">Avec la [co-création](co-authoring-in-excel-add-ins.md), plusieurs personnes peuvent travailler ensemble et modifier le même classeur Excel simultanément. Pour les événements pouvant être déclenchés par un co-auteur, tels que `onChanged`, l’objet **Event** correspondant contient une propriété **source** qui indique si l’événement a été déclenché localement par l’utilisateur actuel (`event.source = Local`) ou par le co-auteur à distance (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="420f7-p103">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="420f7-139">Inscription d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="420f7-139">Register an event handler</span></span>

<span data-ttu-id="420f7-140">L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onChanged` dans la feuille de calcul **Sample**.</span><span class="sxs-lookup"><span data-stu-id="420f7-140">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**.</span></span> <span data-ttu-id="420f7-141">Le code indique que la fonction `handleDataChange` doit être exécutée lorsque les données de la feuille de calcul sont modifiées.</span><span class="sxs-lookup"><span data-stu-id="420f7-141">The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="420f7-142">Gestion d’un événement</span><span class="sxs-lookup"><span data-stu-id="420f7-142">Handle an event</span></span>

<span data-ttu-id="420f7-143">Comme indiqué dans l’exemple précédent, lorsque vous inscrivez un gestionnaire d’événements, vous indiquez la fonction devant être exécutée lorsque l’événement spécifié se produit.</span><span class="sxs-lookup"><span data-stu-id="420f7-143">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs.</span></span> <span data-ttu-id="420f7-144">Vous pouvez créer cette fonction pour effectuer n’importe quelle action nécessaire à votre scénario.</span><span class="sxs-lookup"><span data-stu-id="420f7-144">You can design that function to perform whatever actions your scenario requires.</span></span> <span data-ttu-id="420f7-145">L’exemple de code suivant montre une fonction de gestionnaire d’événements qui écrit simplement des informations sur l’événement dans la console.</span><span class="sxs-lookup"><span data-stu-id="420f7-145">The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="420f7-146">Suppression d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="420f7-146">Remove an event handler</span></span>

<span data-ttu-id="420f7-147">L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onSelectionChanged` dans la feuille de calcul **Sample** et définit la fonction `handleSelectionChange` qui est exécutée lorsqu’un événement se produit.</span><span class="sxs-lookup"><span data-stu-id="420f7-147">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="420f7-148">Il définit également la fonction `remove()` pouvant être appelée par la suite pour supprimer ce gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="420f7-148">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="420f7-149">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="420f7-149">See also</span></span>

- [<span data-ttu-id="420f7-150">Concepts de base de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="420f7-150">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="420f7-151">Spécification libre de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="420f7-151">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)