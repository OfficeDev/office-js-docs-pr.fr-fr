---
title: Utilisation d’événements à l’aide de l’API JavaScript pour Excel
description: ''
ms.date: 01/29/2018
ms.openlocfilehash: 4e04b31e7a130f21d6a9c94d041dc2a122a5890e
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437471"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="af1e5-102">Utilisation d’événements à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="af1e5-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="af1e5-103">Cet article décrit des concepts importants relatifs à l’utilisation des événements dans Excel et fournit des exemples de code montrant comment inscrire des gestionnaires d’événements, gérer des événements et supprimer des gestionnaires d’événements à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="af1e5-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

> [!IMPORTANT]
> <span data-ttu-id="af1e5-104">Les API décrites dans cet article sont actuellement disponibles uniquement dans la version d’évaluation publique (bêta) et ne sont pas destinées à être utilisées dans des environnements de production.</span><span class="sxs-lookup"><span data-stu-id="af1e5-104">The APIs described in this article are currently available only in public preview (beta) and are not intended for use in production environments.</span></span> <span data-ttu-id="af1e5-105">Pour exécuter les exemples de code contenus dans cet article, vous devez utiliser une version suffisamment récente d’Office et faire référence à la bibliothèque bêta du CDN Office.js : https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="af1e5-105">To run the code samples that this article contains, you must use a sufficiently recent build of Office and reference the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

## <a name="events-in-excel"></a><span data-ttu-id="af1e5-106">Événements dans Excel</span><span class="sxs-lookup"><span data-stu-id="af1e5-106">Events in Excel</span></span>

<span data-ttu-id="af1e5-107">Chaque fois que certains types de modifications se produisent dans un classeur Excel, une notification d’événement se déclenche.</span><span class="sxs-lookup"><span data-stu-id="af1e5-107">Each time certain types of changes occur in an Excel workbook, an event notification fires.</span></span> <span data-ttu-id="af1e5-108">En utilisant l’API JavaScript pour Excel, vous pouvez inscrire les gestionnaires d’événements autorisant votre complément à exécuter automatiquement une fonction désignée lorsqu’un événement spécifique se produit.</span><span class="sxs-lookup"><span data-stu-id="af1e5-108">By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs.</span></span> <span data-ttu-id="af1e5-109">Les événements suivants sont actuellement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="af1e5-109">The following events are currently supported.</span></span>

| <span data-ttu-id="af1e5-110">Événement</span><span class="sxs-lookup"><span data-stu-id="af1e5-110">Event</span></span> | <span data-ttu-id="af1e5-111">Description</span><span class="sxs-lookup"><span data-stu-id="af1e5-111">Description</span></span> | <span data-ttu-id="af1e5-112">Objets pris en charge</span><span class="sxs-lookup"><span data-stu-id="af1e5-112">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="af1e5-113">Événement se produisant lors de l’ajout d’un objet.</span><span class="sxs-lookup"><span data-stu-id="af1e5-113">Event that occurs when an object is added.</span></span> | [<span data-ttu-id="af1e5-114">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="af1e5-114">**WorksheetCollection**</span></span>](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetaddedeventargs.md) |
| `onActivated` | <span data-ttu-id="af1e5-115">Événement se produisant lorsqu’un objet est activé.</span><span class="sxs-lookup"><span data-stu-id="af1e5-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="af1e5-116">[**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetactivatedeventargs.md), [**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetactivatedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="af1e5-116">**WorksheetCollection**, [Worksheet](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetactivatedeventargs.md)</span></span> |
| `onDeactivated` | <span data-ttu-id="af1e5-117">Événement se produisant lorsqu’un objet est désactivé.</span><span class="sxs-lookup"><span data-stu-id="af1e5-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="af1e5-118">[**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetdeactivatedeventargs.md), [**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetdeactivatedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="af1e5-118">**WorksheetCollection**, [Worksheet](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetdeactivatedeventargs.md)</span></span> |
| `onChanged` | <span data-ttu-id="af1e5-119">Événement se produisant lorsque les données au sein des cellules sont modifiées.</span><span class="sxs-lookup"><span data-stu-id="af1e5-119">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="af1e5-120">[**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetchangedeventargs.md), [**Table**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tablechangedeventargs.md), [**TableCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tablechangedeventargs.md), [**Liaison**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/bindingdatachangedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="af1e5-120">**Worksheet**, [Table](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetchangedeventargs.md), **TableCollection**, [Binding](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tablechangedeventargs.md)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="af1e5-121">Événement se produisant lorsque la cellule active ou la plage sélectionnée est modifiée.</span><span class="sxs-lookup"><span data-stu-id="af1e5-121">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="af1e5-122">[**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetselectionchangedeventargs.md), [**Table**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tableselectionchangedeventargs.md), [**Liaison**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/bindingselectionchangedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="af1e5-122">**Worksheet**, [Table](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetselectionchangedeventargs.md), **Binding**</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="af1e5-123">Déclencheurs d’événements</span><span class="sxs-lookup"><span data-stu-id="af1e5-123">Event triggers</span></span>

<span data-ttu-id="af1e5-124">Événements au sein d’un classeur Excel pouvant être déclenchés par :</span><span class="sxs-lookup"><span data-stu-id="af1e5-124">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="af1e5-125">Interaction de l’utilisateur via l’interface utilisateur Excel (IU) modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="af1e5-125">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="af1e5-126">Complément (JavaScript) Office modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="af1e5-126">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="af1e5-127">Complément VBA (macro) modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="af1e5-127">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="af1e5-128">Toute modification conforme aux comportements par défaut d’Excel déclenche les événements correspondants dans un classeur.</span><span class="sxs-lookup"><span data-stu-id="af1e5-128">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="af1e5-129">Cycle de vie d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="af1e5-129">Lifecycle of an event handler</span></span>

<span data-ttu-id="af1e5-p103">Un gestionnaire d’événements est créé lorsqu’un complément inscrit le gestionnaire d’événements et est détruit lorsque le complément désinscrit le gestionnaire d’événements ou que le complément est fermé. Les gestionnaires d’événements ne persistent pas en tant que partie du fichier Excel.</span><span class="sxs-lookup"><span data-stu-id="af1e5-p103">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="af1e5-132">Événements et co-création</span><span class="sxs-lookup"><span data-stu-id="af1e5-132">Events and coauthoring</span></span>

<span data-ttu-id="af1e5-p104">Avec la [co-création](co-authoring-in-excel-add-ins.md), plusieurs personnes peuvent travailler ensemble et modifier le même classeur Excel simultanément. Pour les événements pouvant être déclenchés par un co-auteur, tels que `onChanged`, l’objet **Event** correspondant contient une propriété **source** qui indique si l’événement a été déclenché localement par l’utilisateur actuel (`event.source = Local`) ou par le co-auteur à distance (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="af1e5-p104">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="af1e5-135">Inscription d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="af1e5-135">Register an event handler</span></span>

<span data-ttu-id="af1e5-136">L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onChanged` dans la feuille de calcul **Sample**.</span><span class="sxs-lookup"><span data-stu-id="af1e5-136">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**.</span></span> <span data-ttu-id="af1e5-137">Le code indique que la fonction `handleDataChange` doit être exécutée lorsque les données de la feuille de calcul sont modifiées.</span><span class="sxs-lookup"><span data-stu-id="af1e5-137">The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="af1e5-138">Gestion d’un événement</span><span class="sxs-lookup"><span data-stu-id="af1e5-138">Handle an event</span></span>

<span data-ttu-id="af1e5-139">Comme indiqué dans l’exemple précédent, lorsque vous inscrivez un gestionnaire d’événements, vous indiquez la fonction devant être exécutée lorsque l’événement spécifié se produit.</span><span class="sxs-lookup"><span data-stu-id="af1e5-139">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs.</span></span> <span data-ttu-id="af1e5-140">Vous pouvez créer cette fonction pour effectuer n’importe quelle action nécessaire à votre scénario.</span><span class="sxs-lookup"><span data-stu-id="af1e5-140">You can design that function to perform whatever actions your scenario requires.</span></span> <span data-ttu-id="af1e5-141">L’exemple de code suivant montre une fonction de gestionnaire d’événements qui écrit simplement des informations sur l’événement dans la console.</span><span class="sxs-lookup"><span data-stu-id="af1e5-141">The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="af1e5-142">Suppression d’un gestionnaire d’événements</span><span class="sxs-lookup"><span data-stu-id="af1e5-142">Remove an event handler</span></span>

<span data-ttu-id="af1e5-143">L’exemple de code suivant inscrit un gestionnaire d’événements pour l’événement `onSelectionChanged` dans la feuille de calcul **Sample** et définit la fonction `handleSelectionChange` qui est exécutée lorsqu’un événement se produit.</span><span class="sxs-lookup"><span data-stu-id="af1e5-143">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="af1e5-144">Il définit également la fonction `remove()` pouvant être appelée par la suite pour supprimer ce gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="af1e5-144">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="af1e5-145">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="af1e5-145">See also</span></span>

- [<span data-ttu-id="af1e5-146">Concepts de base de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="af1e5-146">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="af1e5-147">Spécification d’ouverture d’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="af1e5-147">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [<span data-ttu-id="af1e5-148">Présentation des fonctionnalités d’événement Excel (aperçu)</span><span class="sxs-lookup"><span data-stu-id="af1e5-148">Introduction to Excel Event Features (preview)</span></span>](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/Event_README.md)
