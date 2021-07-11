---
title: Résolution des problèmes Excel des modules
description: Découvrez comment résoudre les erreurs de développement dans les Excel de développement.
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: cb622a1805be7bec61168ab37a41709a57075788
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349440"
---
# <a name="troubleshooting-excel-add-ins"></a><span data-ttu-id="7b155-103">Résolution des problèmes Excel des modules</span><span class="sxs-lookup"><span data-stu-id="7b155-103">Troubleshooting Excel Add-ins</span></span>

<span data-ttu-id="7b155-104">Cet article traite des problèmes de résolution propres aux Excel.</span><span class="sxs-lookup"><span data-stu-id="7b155-104">This article discusses troubleshooting issues that are unique to Excel.</span></span> <span data-ttu-id="7b155-105">Utilisez l’outil de commentaires en bas de la page pour suggérer d’autres problèmes qui peuvent être ajoutés à l’article.</span><span class="sxs-lookup"><span data-stu-id="7b155-105">Please use the feedback tool at the bottom of the page to suggest other issues that can be added to the article.</span></span>

## <a name="api-limitations-when-the-active-workbook-switches"></a><span data-ttu-id="7b155-106">Limitations de l’API lorsque le workbook actif bascule</span><span class="sxs-lookup"><span data-stu-id="7b155-106">API limitations when the active workbook switches</span></span>

<span data-ttu-id="7b155-107">Les Excel sont destinés à fonctionner sur un seul et même workbook à la fois.</span><span class="sxs-lookup"><span data-stu-id="7b155-107">Add-ins for Excel are intended to operate on a single workbook at a time.</span></span> <span data-ttu-id="7b155-108">Des erreurs peuvent survenir lorsqu’un workbook distinct de celui qui exécute le add-in prend le focus.</span><span class="sxs-lookup"><span data-stu-id="7b155-108">Errors can arise when a workbook that is separate from the one running the add-in gains focus.</span></span> <span data-ttu-id="7b155-109">Cela se produit uniquement lorsque des méthodes particulières sont en cours d’appel lorsque le focus change.</span><span class="sxs-lookup"><span data-stu-id="7b155-109">This only happens when particular methods are in the process of being called when the focus changes.</span></span>

<span data-ttu-id="7b155-110">Les API suivantes sont affectées par ce commutateur de workbook.</span><span class="sxs-lookup"><span data-stu-id="7b155-110">The following APIs are affected by this workbook switch.</span></span>

|<span data-ttu-id="7b155-111">sur les API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="7b155-111">Excel JavaScript API</span></span> | <span data-ttu-id="7b155-112">Erreur lancée</span><span class="sxs-lookup"><span data-stu-id="7b155-112">Error thrown</span></span> |
|--|--|
| `Chart.activate` | <span data-ttu-id="7b155-113">GeneralException</span><span class="sxs-lookup"><span data-stu-id="7b155-113">GeneralException</span></span> |
| `Range.select` | <span data-ttu-id="7b155-114">GeneralException</span><span class="sxs-lookup"><span data-stu-id="7b155-114">GeneralException</span></span> |
| `Table.clearFilters` | <span data-ttu-id="7b155-115">GeneralException</span><span class="sxs-lookup"><span data-stu-id="7b155-115">GeneralException</span></span> |
| `Workbook.getActiveCell`  | <span data-ttu-id="7b155-116">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="7b155-116">InvalidSelection</span></span>|
| `Workbook.getSelectedRange` | <span data-ttu-id="7b155-117">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="7b155-117">InvalidSelection</span></span>|
| `Workbook.getSelectedRanges`  | <span data-ttu-id="7b155-118">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="7b155-118">InvalidSelection</span></span>|
| `Worksheet.activate` | <span data-ttu-id="7b155-119">GeneralException</span><span class="sxs-lookup"><span data-stu-id="7b155-119">GeneralException</span></span> |
| `Worksheet.delete`  | <span data-ttu-id="7b155-120">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="7b155-120">InvalidSelection</span></span>|
| `Worksheet.gridlines` | <span data-ttu-id="7b155-121">GeneralException</span><span class="sxs-lookup"><span data-stu-id="7b155-121">GeneralException</span></span> |
| `Worksheet.showHeadings` | <span data-ttu-id="7b155-122">GeneralException</span><span class="sxs-lookup"><span data-stu-id="7b155-122">GeneralException</span></span> |
| `WorksheetCollection.add` | <span data-ttu-id="7b155-123">GeneralException</span><span class="sxs-lookup"><span data-stu-id="7b155-123">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeAt` | <span data-ttu-id="7b155-124">GeneralException</span><span class="sxs-lookup"><span data-stu-id="7b155-124">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeColumns` | <span data-ttu-id="7b155-125">GeneralException</span><span class="sxs-lookup"><span data-stu-id="7b155-125">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeRows` | <span data-ttu-id="7b155-126">GeneralException</span><span class="sxs-lookup"><span data-stu-id="7b155-126">GeneralException</span></span> |
| `WorksheetFreezePanes.getLocationOrNullObject`| <span data-ttu-id="7b155-127">GeneralException</span><span class="sxs-lookup"><span data-stu-id="7b155-127">GeneralException</span></span> |
| `WorksheetFreezePanes.unfreeze` | <span data-ttu-id="7b155-128">GeneralException</span><span class="sxs-lookup"><span data-stu-id="7b155-128">GeneralException</span></span> |

> [!NOTE]
> <span data-ttu-id="7b155-129">Cela s’applique uniquement à plusieurs Excel de travail ouverts sur Windows ou Mac.</span><span class="sxs-lookup"><span data-stu-id="7b155-129">This only applies to multiple Excel workbooks open on Windows or Mac.</span></span>

## <a name="coauthoring"></a><span data-ttu-id="7b155-130">Co-édition</span><span class="sxs-lookup"><span data-stu-id="7b155-130">Coauthoring</span></span>

<span data-ttu-id="7b155-131">Voir [Co-auteur dans Excel pour](co-authoring-in-excel-add-ins.md) les modèles à utiliser avec des événements dans un environnement de co-auteur.</span><span class="sxs-lookup"><span data-stu-id="7b155-131">See [Coauthoring in Excel add-ins](co-authoring-in-excel-add-ins.md) for patterns to use with events in a coauthoring environment.</span></span> <span data-ttu-id="7b155-132">L’article traite également des conflits potentiels de fusion lors de l’utilisation de certaines API, telles que [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) .</span><span class="sxs-lookup"><span data-stu-id="7b155-132">The article also discusses potential merge conflicts when using certain APIs, such as [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-).</span></span>

## <a name="known-issues"></a><span data-ttu-id="7b155-133">Problèmes connus</span><span class="sxs-lookup"><span data-stu-id="7b155-133">Known Issues</span></span>

### <a name="binding-events-return-temporary-binding-obects"></a><span data-ttu-id="7b155-134">Les événements de liaison retournent `Binding` desobects temporaires</span><span class="sxs-lookup"><span data-stu-id="7b155-134">Binding events return temporary `Binding` obects</span></span>

<span data-ttu-id="7b155-135">[BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding) et [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding) retournent tous deux un objet temporaire qui contient l’ID de l’objet qui a élevé l’événement. `Binding` `Binding`</span><span class="sxs-lookup"><span data-stu-id="7b155-135">Both [BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding) and [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding) return a temporary `Binding` object that contains the ID of the `Binding` object that raised the event.</span></span> <span data-ttu-id="7b155-136">Utilisez cet ID pour `BindingCollection.getItem(id)` récupérer `Binding` l’objet qui a levé l’événement.</span><span class="sxs-lookup"><span data-stu-id="7b155-136">Use this ID with `BindingCollection.getItem(id)` to retrieve the `Binding` object that raised the event.</span></span>

<span data-ttu-id="7b155-137">L’exemple de code suivant montre comment utiliser cet ID de liaison temporaire pour récupérer l’objet `Binding` associé.</span><span class="sxs-lookup"><span data-stu-id="7b155-137">The following code sample shows how to use this temporary binding ID to retrieve the related `Binding` object.</span></span> <span data-ttu-id="7b155-138">Dans l’exemple, un listener d’événement est affecté à une liaison.</span><span class="sxs-lookup"><span data-stu-id="7b155-138">In the sample, an event listener is assigned to a binding.</span></span> <span data-ttu-id="7b155-139">L’écouteur appelle `getBindingId` la méthode lorsque `onDataChanged` l’événement est déclenché.</span><span class="sxs-lookup"><span data-stu-id="7b155-139">The listener calls the `getBindingId` method when the `onDataChanged` event is triggered.</span></span> <span data-ttu-id="7b155-140">La `getBindingId` méthode utilise l’ID de l’objet temporaire pour récupérer `Binding` `Binding` l’objet qui a levé l’événement.</span><span class="sxs-lookup"><span data-stu-id="7b155-140">The `getBindingId` method uses the ID of the temporary `Binding` object to retrieve the `Binding` object that raised the event.</span></span>

```js
Excel.run(function (context) {
    // Retrieve your binding.
    var binding = context.workbook.bindings.getItemAt(0);

    return context.sync().then(function () {
        // Register an event listener to detect changes to your binding
        // and then trigger the `getBindingId` method when the data changes. 
        binding.onDataChanged.add(getBindingId);

        return context.sync();
    });
});

function getBindingId(eventArgs) {
    return Excel.run(function (context) {
        // Get the temporary binding object and load its ID. 
        var tempBindingObject = eventArgs.binding;
        tempBindingObject.load("id");

        // Use the temporary binding object's ID to retrieve the original binding object. 
        var originalBindingObject = context.workbook.bindings.getItem(tempBindingObject.id);

        // You now have the binding object that raised the event: `originalBindingObject`. 
    });
}
```

### <a name="cell-format-usestandardheight-and-usestandardwidth-issues"></a><span data-ttu-id="7b155-141">Format des `useStandardHeight` cellules `useStandardWidth` et problèmes</span><span class="sxs-lookup"><span data-stu-id="7b155-141">Cell format `useStandardHeight` and `useStandardWidth` issues</span></span>

<span data-ttu-id="7b155-142">La [propriété useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) de ne fonctionne pas correctement dans `CellPropertiesFormat` Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="7b155-142">The [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) property of `CellPropertiesFormat` doesn't work properly in Excel on the web.</span></span> <span data-ttu-id="7b155-143">En raison d’un problème dans l Excel sur le Web’interface utilisateur, la définition de la propriété pour calculer la hauteur de manière `useStandardHeight` `true` imprécise sur cette plateforme.</span><span class="sxs-lookup"><span data-stu-id="7b155-143">Due to an issue in the Excel on the web UI, setting the `useStandardHeight` property to `true` calculates height imprecisely on this platform.</span></span> <span data-ttu-id="7b155-144">Par exemple, une hauteur standard de **14** est modifiée à **14,25** Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="7b155-144">For example, a standard height of **14** is modified to **14.25** in Excel on the web.</span></span>

<span data-ttu-id="7b155-145">Sur toutes les plateformes, les propriétés [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) et [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) sont uniquement destinées `CellPropertiesFormat` à être définies sur `true` .</span><span class="sxs-lookup"><span data-stu-id="7b155-145">On all platforms, the [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) and [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) properties of `CellPropertiesFormat` are only intended to be set to `true`.</span></span> <span data-ttu-id="7b155-146">La définition de ces `false` propriétés n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="7b155-146">Setting these properties to `false` has no effect.</span></span> 

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a><span data-ttu-id="7b155-147">Méthode `getImage` Range non pris en Excel pour Mac</span><span class="sxs-lookup"><span data-stu-id="7b155-147">Range `getImage` method unsupported on Excel for Mac</span></span>

<span data-ttu-id="7b155-148">La méthode [Range getImage](/javascript/api/excel/excel.range#getImage__) n’est actuellement pas prise en charge dans Excel pour Mac.</span><span class="sxs-lookup"><span data-stu-id="7b155-148">The Range [getImage](/javascript/api/excel/excel.range#getImage__) method isn't currently supported in Excel for Mac.</span></span> <span data-ttu-id="7b155-149">Consultez [la #235 OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues/235) pour l’état actuel.</span><span class="sxs-lookup"><span data-stu-id="7b155-149">See [OfficeDev/office-js Issue #235](https://github.com/OfficeDev/office-js/issues/235) for the current status.</span></span>

### <a name="range-return-character-limit"></a><span data-ttu-id="7b155-150">Limite de caractères de retour de plage</span><span class="sxs-lookup"><span data-stu-id="7b155-150">Range return character limit</span></span>

<span data-ttu-id="7b155-151">Les [méthodes Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#getRange_address_) et [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#getRanges_address_) ont une limite de chaîne d’adresses de 8 192 caractères.</span><span class="sxs-lookup"><span data-stu-id="7b155-151">The [Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#getRange_address_) and [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#getRanges_address_) methods have an address string limit of 8192 characters.</span></span> <span data-ttu-id="7b155-152">Lorsque cette limite est dépassée, la chaîne d’adresse est tronquée à 8 192 caractères.</span><span class="sxs-lookup"><span data-stu-id="7b155-152">When this limit is exceeded, the address string is truncated to 8192 characters.</span></span>

## <a name="see-also"></a><span data-ttu-id="7b155-153">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="7b155-153">See also</span></span>

- [<span data-ttu-id="7b155-154">Résoudre les erreurs de développement avec Office de recherche</span><span class="sxs-lookup"><span data-stu-id="7b155-154">Troubleshoot development errors with Office Add-ins</span></span>](../testing/troubleshoot-development-errors.md)
- [<span data-ttu-id="7b155-155">Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office</span><span class="sxs-lookup"><span data-stu-id="7b155-155">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
