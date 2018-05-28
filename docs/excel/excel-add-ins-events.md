---
title: Utilisation d??v?nements ? l?aide de l?API JavaScript pour Excel
description: ''
ms.date: 01/29/2018
ms.openlocfilehash: 4e04b31e7a130f21d6a9c94d041dc2a122a5890e
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="work-with-events-using-the-excel-javascript-api"></a><span data-ttu-id="6f7b9-102">Utilisation d??v?nements ? l?aide de l?API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="6f7b9-102">Work with Events using the Excel JavaScript API</span></span> 

<span data-ttu-id="6f7b9-103">Cet article d?crit des concepts importants relatifs ? l?utilisation des ?v?nements dans Excel et fournit des exemples de code montrant comment inscrire des gestionnaires d??v?nements, g?rer des ?v?nements et supprimer des gestionnaires d??v?nements ? l?aide de l?API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-103">This article describes important concepts related to working with events in Excel and provides code samples that show how to register event handlers, handle events, and remove event handlers using the Excel JavaScript API.</span></span> 

> [!IMPORTANT]
> <span data-ttu-id="6f7b9-104">Les API d?crites dans cet article sont actuellement disponibles uniquement dans la version d??valuation publique (b?ta) et ne sont pas destin?es ? ?tre utilis?es dans des environnements de production.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-104">The APIs described in this article are currently available only in public preview (beta) and are not intended for use in production environments.</span></span> <span data-ttu-id="6f7b9-105">Pour ex?cuter les exemples de code contenus dans cet article, vous devez utiliser une version suffisamment r?cente d?Office et faire r?f?rence ? la biblioth?que b?ta du CDN Office.js : https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-105">To run the code samples that this article contains, you must use a sufficiently recent build of Office and reference the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

## <a name="events-in-excel"></a><span data-ttu-id="6f7b9-106">?v?nements dans Excel</span><span class="sxs-lookup"><span data-stu-id="6f7b9-106">Events in Excel</span></span>

<span data-ttu-id="6f7b9-107">Chaque fois que certains types de modifications se produisent dans un classeur Excel, une notification d??v?nement se d?clenche.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-107">Each time certain types of changes occur in an Excel workbook, an event notification fires.</span></span> <span data-ttu-id="6f7b9-108">En utilisant l?API JavaScript pour Excel, vous pouvez inscrire les gestionnaires d??v?nements autorisant votre compl?ment ? ex?cuter automatiquement une fonction d?sign?e lorsqu?un ?v?nement sp?cifique se produit.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-108">By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs.</span></span> <span data-ttu-id="6f7b9-109">Les ?v?nements suivants sont actuellement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-109">The following events are currently supported.</span></span>

| <span data-ttu-id="6f7b9-110">?v?nement</span><span class="sxs-lookup"><span data-stu-id="6f7b9-110">Event</span></span> | <span data-ttu-id="6f7b9-111">Description</span><span class="sxs-lookup"><span data-stu-id="6f7b9-111">Description</span></span> | <span data-ttu-id="6f7b9-112">Objets pris en charge</span><span class="sxs-lookup"><span data-stu-id="6f7b9-112">Supported objects</span></span> |
|:---------------|:-------------|:-----------|
| `onAdded` | <span data-ttu-id="6f7b9-113">?v?nement se produisant lors de l?ajout d?un objet.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-113">Event that occurs when an object is added.</span></span> | [<span data-ttu-id="6f7b9-114">**WorksheetCollection**</span><span class="sxs-lookup"><span data-stu-id="6f7b9-114">**WorksheetCollection**</span></span>](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetaddedeventargs.md) |
| `onActivated` | <span data-ttu-id="6f7b9-115">?v?nement se produisant lorsqu?un objet est activ?.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-115">Event that occurs when an object is activated.</span></span> | <span data-ttu-id="6f7b9-116">[**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetactivatedeventargs.md), [**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetactivatedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="6f7b9-116">**WorksheetCollection**, [Worksheet](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetactivatedeventargs.md)</span></span> |
| `onDeactivated` | <span data-ttu-id="6f7b9-117">?v?nement se produisant lorsqu?un objet est d?sactiv?.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-117">Event that occurs when an object is deactivated.</span></span> | <span data-ttu-id="6f7b9-118">[**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetdeactivatedeventargs.md), [**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetdeactivatedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="6f7b9-118">**WorksheetCollection**, [Worksheet](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetdeactivatedeventargs.md)</span></span> |
| `onChanged` | <span data-ttu-id="6f7b9-119">?v?nement se produisant lorsque les donn?es au sein des cellules sont modifi?es.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-119">Event that occurs when data within cells is changed.</span></span> | <span data-ttu-id="6f7b9-120">[**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetchangedeventargs.md), [**Table**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tablechangedeventargs.md), [**TableCollection**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tablechangedeventargs.md), [**Liaison**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/bindingdatachangedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="6f7b9-120">**Worksheet**, [Table](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetchangedeventargs.md), **TableCollection**, [Binding](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tablechangedeventargs.md)</span></span> |
| `onSelectionChanged` | <span data-ttu-id="6f7b9-121">?v?nement se produisant lorsque la cellule active ou la plage s?lectionn?e est modifi?e.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-121">Event that occurs when the active cell or selected range is changed.</span></span> | <span data-ttu-id="6f7b9-122">[**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetselectionchangedeventargs.md), [**Table**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/tableselectionchangedeventargs.md), [**Liaison**](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/bindingselectionchangedeventargs.md)</span><span class="sxs-lookup"><span data-stu-id="6f7b9-122">**Worksheet**, [Table](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/worksheetselectionchangedeventargs.md), **Binding**</span></span> |

### <a name="event-triggers"></a><span data-ttu-id="6f7b9-123">D?clencheurs d??v?nements</span><span class="sxs-lookup"><span data-stu-id="6f7b9-123">Event triggers</span></span>

<span data-ttu-id="6f7b9-124">?v?nements au sein d?un classeur Excel pouvant ?tre d?clench?s par :</span><span class="sxs-lookup"><span data-stu-id="6f7b9-124">Events within an Excel workbook can be triggered by:</span></span>

- <span data-ttu-id="6f7b9-125">Interaction de l?utilisateur via l?interface utilisateur Excel (IU) modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="6f7b9-125">User interaction via the Excel user interface (UI) that changes the workbook</span></span>
- <span data-ttu-id="6f7b9-126">Compl?ment (JavaScript) Office modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="6f7b9-126">Office add-in (JavaScript) code that changes the workbook</span></span>
- <span data-ttu-id="6f7b9-127">Compl?ment VBA (macro) modifiant le classeur</span><span class="sxs-lookup"><span data-stu-id="6f7b9-127">VBA add-in (macro) code that changes the workbook</span></span>

<span data-ttu-id="6f7b9-128">Toute modification conforme aux comportements par d?faut d?Excel d?clenche les ?v?nements correspondants dans un classeur.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-128">Any change that complies with default behavior of Excel will trigger the corresponding event(s) in a workbook.</span></span>

### <a name="lifecycle-of-an-event-handler"></a><span data-ttu-id="6f7b9-129">Cycle de vie d?un gestionnaire d??v?nements</span><span class="sxs-lookup"><span data-stu-id="6f7b9-129">Lifecycle of an event handler</span></span>

<span data-ttu-id="6f7b9-p103">Un gestionnaire d??v?nements est cr?? lorsqu?un compl?ment inscrit le gestionnaire d??v?nements et est d?truit lorsque le compl?ment d?sinscrit le gestionnaire d??v?nements ou que le compl?ment est ferm?. Les gestionnaires d??v?nements ne persistent pas en tant que partie du fichier Excel.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-p103">An event handler is created when an add-in registers the event handler and is destroyed when the add-in unregisters the event handler or when the add-in is closed. Event handlers do not persist as part of the Excel file.</span></span>

### <a name="events-and-coauthoring"></a><span data-ttu-id="6f7b9-132">?v?nements et co-cr?ation</span><span class="sxs-lookup"><span data-stu-id="6f7b9-132">Events and coauthoring</span></span>

<span data-ttu-id="6f7b9-p104">Avec la [co-cr?ation](co-authoring-in-excel-add-ins.md), plusieurs personnes peuvent travailler ensemble et modifier le m?me classeur Excel simultan?ment. Pour les ?v?nements pouvant ?tre d?clench?s par un co-auteur, tels que `onChanged`, l?objet **Event** correspondant contient une propri?t? **source** qui indique si l??v?nement a ?t? d?clench? localement par l?utilisateur actuel (`event.source = Local`) ou par le co-auteur ? distance (`event.source = Remote`).</span><span class="sxs-lookup"><span data-stu-id="6f7b9-p104">With [coauthoring](co-authoring-in-excel-add-ins.md), multiple people can work together and edit the same Excel workbook simultaneously. For events that can be triggered by a coauthor, such as `onChanged`, the corresponding **Event** object will contain a **source** property that indicates whether the event was triggered locally by the current user (`event.source = Local`) or was triggered by the remote coauthor (`event.source = Remote`).</span></span>

## <a name="register-an-event-handler"></a><span data-ttu-id="6f7b9-135">Inscription d?un gestionnaire d??v?nements</span><span class="sxs-lookup"><span data-stu-id="6f7b9-135">Register an event handler</span></span>

<span data-ttu-id="6f7b9-136">L?exemple de code suivant inscrit un gestionnaire d??v?nements pour l??v?nement `onChanged` dans la feuille de calcul **Sample**.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-136">The following code sample registers an event handler for the `onChanged` event in the worksheet named **Sample**.</span></span> <span data-ttu-id="6f7b9-137">Le code indique que la fonction `handleDataChange` doit ?tre ex?cut?e lorsque les donn?es de la feuille de calcul sont modifi?es.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-137">The code specifies that when data changes in that worksheet, the `handleDataChange` function should run.</span></span>

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

## <a name="handle-an-event"></a><span data-ttu-id="6f7b9-138">Gestion d?un ?v?nement</span><span class="sxs-lookup"><span data-stu-id="6f7b9-138">Handle an event</span></span>

<span data-ttu-id="6f7b9-139">Comme indiqu? dans l?exemple pr?c?dent, lorsque vous inscrivez un gestionnaire d??v?nements, vous indiquez la fonction devant ?tre ex?cut?e lorsque l??v?nement sp?cifi? se produit.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-139">As shown in the previous example, when you register an event handler, you indicate the function that should run when the specified event occurs.</span></span> <span data-ttu-id="6f7b9-140">Vous pouvez cr?er cette fonction pour effectuer n?importe quelle action n?cessaire ? votre sc?nario.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-140">You can design that function to perform whatever actions your scenario requires.</span></span> <span data-ttu-id="6f7b9-141">L?exemple de code suivant montre une fonction de gestionnaire d??v?nements qui ?crit simplement des informations sur l??v?nement dans la console.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-141">The following code sample shows an event handler function that simply writes information about the event to the console.</span></span> 

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

## <a name="remove-an-event-handler"></a><span data-ttu-id="6f7b9-142">Suppression d?un gestionnaire d??v?nements</span><span class="sxs-lookup"><span data-stu-id="6f7b9-142">Remove an event handler</span></span>

<span data-ttu-id="6f7b9-143">L?exemple de code suivant inscrit un gestionnaire d??v?nements pour l??v?nement `onSelectionChanged` dans la feuille de calcul **Sample** et d?finit la fonction `handleSelectionChange` qui est ex?cut?e lorsqu?un ?v?nement se produit.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-143">The following code sample registers an event handler for the `onSelectionChanged` event in the worksheet named **Sample** and defines the `handleSelectionChange` function that will run when the event occurs.</span></span> <span data-ttu-id="6f7b9-144">Il d?finit ?galement la fonction `remove()` pouvant ?tre appel?e par la suite pour supprimer ce gestionnaire d??v?nements.</span><span class="sxs-lookup"><span data-stu-id="6f7b9-144">It also defines the `remove()` function that can subsequently be called to remove that event handler.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="6f7b9-145">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="6f7b9-145">See also</span></span>

- [<span data-ttu-id="6f7b9-146">Concepts de base de l?API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="6f7b9-146">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="6f7b9-147">Sp?cification d?ouverture d?API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="6f7b9-147">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [<span data-ttu-id="6f7b9-148">Pr?sentation des fonctionnalit?s d??v?nement Excel (aper?u)</span><span class="sxs-lookup"><span data-stu-id="6f7b9-148">Introduction to Excel Event Features (preview)</span></span>](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/Event_README.md)
