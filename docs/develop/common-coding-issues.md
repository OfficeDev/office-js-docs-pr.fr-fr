---
title: Conseils de codage pour les problèmes courants et les comportements de plateforme inattendus
description: Liste des problèmes de plateforme d’API JavaScript pour Office fréquemment rencontrés par les développeurs.
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: f6d6a31059b32550e3176ed278d7da4c2c7a6c68
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292910"
---
# <a name="coding-guidance-for-common-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="cc054-103">Conseils de codage pour les problèmes courants et les comportements de plateforme inattendus</span><span class="sxs-lookup"><span data-stu-id="cc054-103">Coding guidance for common issues and unexpected platform behaviors</span></span>

<span data-ttu-id="cc054-104">Cet article met en évidence les aspects de l’API JavaScript pour Office qui peuvent entraîner un comportement inattendu ou nécessiter des modèles de codage spécifiques pour obtenir le résultat souhaité.</span><span class="sxs-lookup"><span data-stu-id="cc054-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="cc054-105">Si vous rencontrez un problème qui se trouve dans cette liste, faites-le nous connaître en utilisant le formulaire de commentaires au bas de l’article.</span><span class="sxs-lookup"><span data-stu-id="cc054-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a><span data-ttu-id="cc054-106">Les API communes et les API Outlook ne sont pas basées sur la promesse</span><span class="sxs-lookup"><span data-stu-id="cc054-106">Common APIs and Outlook APIs are not promise-based</span></span>

<span data-ttu-id="cc054-107">Les [API communes](/javascript/api/office) (celles qui ne sont pas liées à une application Office particulière) et les [API Outlook](/javascript/api/outlook) utilisent un modèle de programmation basé sur les rappels.</span><span class="sxs-lookup"><span data-stu-id="cc054-107">The [Common APIs](/javascript/api/office) (those that are not tied to a particular Office application) and [Outlook APIs](/javascript/api/outlook) use a callback-based programming model.</span></span> <span data-ttu-id="cc054-108">L’interaction avec le document Office sous-jacent nécessite un appel asynchrone en lecture ou en écriture qui spécifie un rappel à exécuter lorsque l’opération se termine.</span><span class="sxs-lookup"><span data-stu-id="cc054-108">Interacting with the underlying Office document requires an asynchronous read or write call that specifies a callback to be run when the operation completes.</span></span> <span data-ttu-id="cc054-109">Pour obtenir un exemple de ce modèle, consultez la rubrique [document. getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="cc054-109">For an example of this pattern, see [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span>

<span data-ttu-id="cc054-110">Ces méthodes d’API et d’API courantes ne renvoient pas de [promesses](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span><span class="sxs-lookup"><span data-stu-id="cc054-110">These Common API and Outlook API methods do not return [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span></span> <span data-ttu-id="cc054-111">Par conséquent, vous ne pouvez pas utiliser [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) pour suspendre l’exécution jusqu’à la fin de l’opération asynchrone.</span><span class="sxs-lookup"><span data-stu-id="cc054-111">Therefore, you cannot use [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) to pause the execution until the asynchronous operation completes.</span></span> <span data-ttu-id="cc054-112">Si vous avez besoin `await` de comportement, vous pouvez encapsuler l’appel de méthode dans une promesse créée de manière explicite.</span><span class="sxs-lookup"><span data-stu-id="cc054-112">If you need `await` behavior, you can wrap the method call in an explicitly created Promise.</span></span>

```js
readDocumentFileAsync(): Promise<any> {
    return new Promise((resolve, reject) => {
        const chunkSize = 65536;
        const self = this;

        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: chunkSize }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(asyncResult.error);
            } else {
                // `getAllSlices` is a Promise-wrapped implementation of File.getSliceAsync.
                self.getAllSlices(asyncResult.value).then(result => {
                    if (result.IsSuccess) {
                        resolve(result.Data);
                    } else {
                        reject(asyncResult.error);
                    }
                });
            }
        });
    });
}
```

> [!NOTE]
> <span data-ttu-id="cc054-113">La documentation de référence contient l’implémentation encapsulée de [fichier. getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span><span class="sxs-lookup"><span data-stu-id="cc054-113">The reference documentation contains the Promise-wrapped implementation of [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span></span>

## <a name="some-properties-cannot-be-set-directly"></a><span data-ttu-id="cc054-114">Certaines propriétés ne peuvent pas être définies directement</span><span class="sxs-lookup"><span data-stu-id="cc054-114">Some properties cannot be set directly</span></span>

> [!NOTE]
> <span data-ttu-id="cc054-115">Cette section s’applique uniquement aux API propres à l’application pour Excel et Word.</span><span class="sxs-lookup"><span data-stu-id="cc054-115">This section only applies to the application-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="cc054-116">Certaines propriétés ne peuvent pas être définies, bien qu’elles soient accessibles en écriture.</span><span class="sxs-lookup"><span data-stu-id="cc054-116">Some properties cannot be set, despite being writable.</span></span> <span data-ttu-id="cc054-117">Ces propriétés font partie d’une propriété parent qui doit être définie en tant qu’objet unique.</span><span class="sxs-lookup"><span data-stu-id="cc054-117">These properties are part of a parent property that must be set as a single object.</span></span> <span data-ttu-id="cc054-118">Cela est dû au fait que cette propriété Parent repose sur les sous-propriétés ayant des relations logiques spécifiques.</span><span class="sxs-lookup"><span data-stu-id="cc054-118">This is because that parent property relies on the subproperties having specific, logical relationships.</span></span> <span data-ttu-id="cc054-119">Ces propriétés parent doivent être définies à l’aide de la notation littérale d’objet pour définir l’objet entier, au lieu de définir les sous-propriétés individuelles de cet objet.</span><span class="sxs-lookup"><span data-stu-id="cc054-119">These parent properties must be set using object literal notation to set the entire object, instead of setting that object's individual subproperties.</span></span> <span data-ttu-id="cc054-120">Vous trouverez un exemple dans [PageLayout](/javascript/api/excel/excel.pagelayout).</span><span class="sxs-lookup"><span data-stu-id="cc054-120">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="cc054-121">La `zoom` propriété doit être définie avec un seul objet [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) , comme illustré ci-dessous :</span><span class="sxs-lookup"><span data-stu-id="cc054-121">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="cc054-122">Dans l’exemple précédent, vous ne seriez ***pas*** en mesure d’affecter directement `zoom` une valeur : `sheet.pageLayout.zoom.scale = 200;` .</span><span class="sxs-lookup"><span data-stu-id="cc054-122">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="cc054-123">Cette instruction génère une erreur car `zoom` elle n’est pas chargée.</span><span class="sxs-lookup"><span data-stu-id="cc054-123">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="cc054-124">Même si `zoom` elles ont été chargées, l’ensemble de l’étendue ne prendra pas effet.</span><span class="sxs-lookup"><span data-stu-id="cc054-124">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="cc054-125">Toutes les opérations de contexte se produisent `zoom` , actualisant l’objet proxy dans le complément et remplaçant les valeurs définies localement.</span><span class="sxs-lookup"><span data-stu-id="cc054-125">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="cc054-126">Ce comportement diffère des [Propriétés de navigation](application-specific-api-model.md#scalar-and-navigation-properties) telles que [Range. format](/javascript/api/excel/excel.range#format).</span><span class="sxs-lookup"><span data-stu-id="cc054-126">This behavior differs from [navigational properties](application-specific-api-model.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="cc054-127">Les propriétés de `format` peuvent être définies à l’aide de la navigation d’objet, comme illustré ci-dessous :</span><span class="sxs-lookup"><span data-stu-id="cc054-127">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="cc054-128">Vous pouvez identifier une propriété qui ne peut pas avoir ses sous-propriétés directement définies en vérifiant son modificateur en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="cc054-128">You can identify a property that cannot have its subproperties directly set by checking its read-only modifier.</span></span> <span data-ttu-id="cc054-129">Les propriétés non en lecture seule de toutes les propriétés en lecture seule peuvent être définies directement.</span><span class="sxs-lookup"><span data-stu-id="cc054-129">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="cc054-130">Les propriétés accessibles en écriture comme `PageLayout.zoom` doivent être définies avec un objet à ce niveau.</span><span class="sxs-lookup"><span data-stu-id="cc054-130">Writeable properties like `PageLayout.zoom` must be set with an object at that level.</span></span> <span data-ttu-id="cc054-131">En Résumé :</span><span class="sxs-lookup"><span data-stu-id="cc054-131">In summary:</span></span>

- <span data-ttu-id="cc054-132">Propriété en lecture seule : les sous-propriétés peuvent être définies via la navigation.</span><span class="sxs-lookup"><span data-stu-id="cc054-132">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="cc054-133">Propriété accessible en écriture : les sous-propriétés ne peuvent pas être définies par le biais de la navigation (elles doivent être définies dans le cadre de l’attribution initiale de l’objet parent).</span><span class="sxs-lookup"><span data-stu-id="cc054-133">Writable property: Subproperties cannot be set through navigation (must be set as part of the initial parent object assignment).</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="cc054-134">Définition de propriétés en lecture seule</span><span class="sxs-lookup"><span data-stu-id="cc054-134">Setting read-only properties</span></span>

<span data-ttu-id="cc054-135">Les [définitions](referencing-the-javascript-api-for-office-library-from-its-cdn.md) de la machine à écrire pour Office js spécifient les propriétés d’objet en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="cc054-135">The [TypeScript definitions](referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="cc054-136">Si vous tentez de définir une propriété en lecture seule, l’opération d’écriture échoue sans avertissement, sans qu’aucune erreur ne soit générée.</span><span class="sxs-lookup"><span data-stu-id="cc054-136">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="cc054-137">L’exemple suivant tente à tort de définir la propriété en lecture seule [Chart.ID](/javascript/api/excel/excel.chart#id).</span><span class="sxs-lookup"><span data-stu-id="cc054-137">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="removing-event-handlers"></a><span data-ttu-id="cc054-138">Suppression de gestionnaires d’événements</span><span class="sxs-lookup"><span data-stu-id="cc054-138">Removing event handlers</span></span>

<span data-ttu-id="cc054-139">Les gestionnaires d’événements doivent être supprimés à l’aide du même `RequestContext` que celui dans lequel ils ont été ajoutés.</span><span class="sxs-lookup"><span data-stu-id="cc054-139">Event handlers must be removed using the same `RequestContext` in which they were added.</span></span> <span data-ttu-id="cc054-140">Si vous avez besoin que votre complément supprime un gestionnaire d’événements en cours d’exécution, vous devez stocker l’objet Context utilisé pour ajouter le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="cc054-140">If you need your add-in to remove an event handler while running, you'll need to store the context object used to add the handler.</span></span>

```js
Excel.run(async (context) => {
    [...]

    // To later remove an event handler, store the context somewhere accessible to the handler removal function.
    // You may find it helpful to also store the event handler object and associate it with the context.
    selectionChangedHandler = myWorksheet.onSelectionChanged.add(callback);
    savedContext = currentContext;
    return context.sync();
}
```

## <a name="supporting-internet-explorer"></a><span data-ttu-id="cc054-141">Prise en charge d’Internet Explorer</span><span class="sxs-lookup"><span data-stu-id="cc054-141">Supporting Internet Explorer</span></span>

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## <a name="excel-specific-issues"></a><span data-ttu-id="cc054-142">Problèmes spécifiques à Excel</span><span class="sxs-lookup"><span data-stu-id="cc054-142">Excel-specific issues</span></span>

### <a name="api-limitations-when-the-active-workbook-switches"></a><span data-ttu-id="cc054-143">Limitations de l’API lorsque le classeur actif bascule</span><span class="sxs-lookup"><span data-stu-id="cc054-143">API limitations when the active workbook switches</span></span>

<span data-ttu-id="cc054-144">Les compléments pour Excel sont conçus pour fonctionner sur un seul classeur à la fois.</span><span class="sxs-lookup"><span data-stu-id="cc054-144">Add-ins for Excel are intended to operate on a single workbook at a time.</span></span> <span data-ttu-id="cc054-145">Des erreurs peuvent se produire lorsqu’un classeur distinct de celui qui exécute le complément obtient le focus.</span><span class="sxs-lookup"><span data-stu-id="cc054-145">Errors can arise when a workbook that is separate from the one running the add-in gains focus.</span></span> <span data-ttu-id="cc054-146">Cela se produit uniquement lorsque des méthodes particulières sont en cours d’appel lorsque le focus est modifié.</span><span class="sxs-lookup"><span data-stu-id="cc054-146">This only happens when particular methods are in the process of being called when the focus changes.</span></span>

<span data-ttu-id="cc054-147">Les API suivantes sont affectées par ce commutateur de classeurs :</span><span class="sxs-lookup"><span data-stu-id="cc054-147">The following APIs are affected by this workbook switch:</span></span>

|<span data-ttu-id="cc054-148">sur les API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="cc054-148">Excel JavaScript API</span></span> | <span data-ttu-id="cc054-149">Erreur générée</span><span class="sxs-lookup"><span data-stu-id="cc054-149">Error thrown</span></span> |
|--|--|
| `Chart.activate` | <span data-ttu-id="cc054-150">GeneralException</span><span class="sxs-lookup"><span data-stu-id="cc054-150">GeneralException</span></span> |
| `Range.select` | <span data-ttu-id="cc054-151">GeneralException</span><span class="sxs-lookup"><span data-stu-id="cc054-151">GeneralException</span></span> |
| `Table.clearFilters` | <span data-ttu-id="cc054-152">GeneralException</span><span class="sxs-lookup"><span data-stu-id="cc054-152">GeneralException</span></span> |
| `Workbook.getActiveCell`  | <span data-ttu-id="cc054-153">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="cc054-153">InvalidSelection</span></span>|
| `Workbook.getSelectedRange` | <span data-ttu-id="cc054-154">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="cc054-154">InvalidSelection</span></span>|
| `Workbook.getSelectedRanges`  | <span data-ttu-id="cc054-155">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="cc054-155">InvalidSelection</span></span>|
| `Worksheet.activate` | <span data-ttu-id="cc054-156">GeneralException</span><span class="sxs-lookup"><span data-stu-id="cc054-156">GeneralException</span></span> |
| `Worksheet.delete`  | <span data-ttu-id="cc054-157">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="cc054-157">InvalidSelection</span></span>|
| `Worksheet.gridlines` | <span data-ttu-id="cc054-158">GeneralException</span><span class="sxs-lookup"><span data-stu-id="cc054-158">GeneralException</span></span> |
| `Worksheet.showHeadings` | <span data-ttu-id="cc054-159">GeneralException</span><span class="sxs-lookup"><span data-stu-id="cc054-159">GeneralException</span></span> |
| `WorksheetCollection.add` | <span data-ttu-id="cc054-160">GeneralException</span><span class="sxs-lookup"><span data-stu-id="cc054-160">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeAt` | <span data-ttu-id="cc054-161">GeneralException</span><span class="sxs-lookup"><span data-stu-id="cc054-161">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeColumns` | <span data-ttu-id="cc054-162">GeneralException</span><span class="sxs-lookup"><span data-stu-id="cc054-162">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeRows` | <span data-ttu-id="cc054-163">GeneralException</span><span class="sxs-lookup"><span data-stu-id="cc054-163">GeneralException</span></span> |
| `WorksheetFreezePanes.getLocationOrNullObject`| <span data-ttu-id="cc054-164">GeneralException</span><span class="sxs-lookup"><span data-stu-id="cc054-164">GeneralException</span></span> |
| `WorksheetFreezePanes.unfreeze` | <span data-ttu-id="cc054-165">GeneralException</span><span class="sxs-lookup"><span data-stu-id="cc054-165">GeneralException</span></span> |

> [!NOTE]
> <span data-ttu-id="cc054-166">Cela s’applique uniquement à plusieurs classeurs Excel ouverts sous Windows ou Mac.</span><span class="sxs-lookup"><span data-stu-id="cc054-166">This only applies to multiple Excel workbooks open on Windows or Mac.</span></span>

### <a name="coauthoring"></a><span data-ttu-id="cc054-167">Co-édition</span><span class="sxs-lookup"><span data-stu-id="cc054-167">Coauthoring</span></span>

<span data-ttu-id="cc054-168">Consultez la rubrique [co-authoring in Excel Add-ins](../excel/co-authoring-in-excel-add-ins.md) for patterns to use with Events in a CoAuthoring Environment.</span><span class="sxs-lookup"><span data-stu-id="cc054-168">See [Coauthoring in Excel add-ins](../excel/co-authoring-in-excel-add-ins.md) for patterns to use with events in a coauthoring environment.</span></span> <span data-ttu-id="cc054-169">L’article aborde également les conflits de fusion potentiels lors de l’utilisation de certaines API, telles que [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) .</span><span class="sxs-lookup"><span data-stu-id="cc054-169">The article also discusses potential merge conflicts when using certain APIs, such as [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-).</span></span>

## <a name="see-also"></a><span data-ttu-id="cc054-170">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="cc054-170">See also</span></span>

- [<span data-ttu-id="cc054-171">Limites des ressources et optimisation des performances pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="cc054-171">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
- <span data-ttu-id="cc054-172">[OfficeDev/Office-js](https://github.com/OfficeDev/office-js/issues): le lieu de signaler et d’afficher les problèmes liés à la plateforme des compléments Office et aux API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="cc054-172">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="cc054-173">[Débordement de pile](https://stackoverflow.com/questions/tagged/office-js): emplacement où poser des questions de programmation sur les API JavaScript Office.</span><span class="sxs-lookup"><span data-stu-id="cc054-173">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="cc054-174">Veillez à appliquer la balise « Office-js » à votre question lors de la publication dans le débordement de pile.</span><span class="sxs-lookup"><span data-stu-id="cc054-174">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="cc054-175">[UserVoice](https://officespdev.uservoice.com/): le lieu de suggérer de nouvelles fonctionnalités pour la plateforme des compléments Office et les API JavaScript pour Office.</span><span class="sxs-lookup"><span data-stu-id="cc054-175">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
