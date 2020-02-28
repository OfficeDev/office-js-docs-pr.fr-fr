---
title: Conseils de codage pour les problèmes courants et les comportements de plateforme inattendus
description: Liste des problèmes de plateforme d’API JavaScript pour Office fréquemment rencontrés par les développeurs.
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: da6986149172238963a06b3296d9fdd7c2411d9d
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324609"
---
# <a name="coding-guidance-for-common-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="c78c9-103">Conseils de codage pour les problèmes courants et les comportements de plateforme inattendus</span><span class="sxs-lookup"><span data-stu-id="c78c9-103">Coding guidance for common issues and unexpected platform behaviors</span></span>

<span data-ttu-id="c78c9-104">Cet article met en évidence les aspects de l’API JavaScript pour Office qui peuvent entraîner un comportement inattendu ou nécessiter des modèles de codage spécifiques pour obtenir le résultat souhaité.</span><span class="sxs-lookup"><span data-stu-id="c78c9-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="c78c9-105">Si vous rencontrez un problème qui se trouve dans cette liste, faites-le nous connaître en utilisant le formulaire de commentaires au bas de l’article.</span><span class="sxs-lookup"><span data-stu-id="c78c9-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a><span data-ttu-id="c78c9-106">Les API communes et les API Outlook ne sont pas basées sur la promesse</span><span class="sxs-lookup"><span data-stu-id="c78c9-106">Common APIs and Outlook APIs are not promise-based</span></span>

<span data-ttu-id="c78c9-107">Les [API communes](/javascript/api/office) (qui ne sont pas liées à un hôte Office particulier) et les [API Outlook](/javascript/api/outlook) utilisent un modèle de programmation basé sur les rappels.</span><span class="sxs-lookup"><span data-stu-id="c78c9-107">The [Common APIs](/javascript/api/office) (those that are not tied to a particular Office host) and [Outlook APIs](/javascript/api/outlook) use a callback-based programming model.</span></span> <span data-ttu-id="c78c9-108">L’interaction avec le document Office sous-jacent nécessite un appel asynchrone en lecture ou en écriture qui spécifie un rappel à exécuter lorsque l’opération se termine.</span><span class="sxs-lookup"><span data-stu-id="c78c9-108">Interacting with the underlying Office document requires an asynchronous read or write call that specifies a callback to be ran when the operation completes.</span></span> <span data-ttu-id="c78c9-109">Pour obtenir un exemple de ce modèle, consultez la rubrique [document. getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="c78c9-109">For an example of this pattern, see [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span>

<span data-ttu-id="c78c9-110">Ces méthodes d’API et d’API courantes ne renvoient pas de [promesses](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span><span class="sxs-lookup"><span data-stu-id="c78c9-110">These Common API and Outlook API methods do not return [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span></span> <span data-ttu-id="c78c9-111">Par conséquent, vous ne pouvez pas utiliser [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) pour suspendre l’exécution jusqu’à la fin de l’opération asynchrone.</span><span class="sxs-lookup"><span data-stu-id="c78c9-111">Therefore, you cannot use [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) to pause the execution until the asynchronous operation completes.</span></span> <span data-ttu-id="c78c9-112">Si vous avez `await` besoin de comportement, vous pouvez encapsuler l’appel de méthode dans une promesse créée de manière explicite.</span><span class="sxs-lookup"><span data-stu-id="c78c9-112">If you need `await` behavior, you can wrap the method call in an explicitly created Promise.</span></span>

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
> <span data-ttu-id="c78c9-113">La documentation de référence contient l’implémentation encapsulée de [fichier. getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span><span class="sxs-lookup"><span data-stu-id="c78c9-113">The reference documentation contains the Promise-wrapped implementation of [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span></span>

## <a name="some-properties-cannot-be-set-directly"></a><span data-ttu-id="c78c9-114">Certaines propriétés ne peuvent pas être définies directement</span><span class="sxs-lookup"><span data-stu-id="c78c9-114">Some properties cannot be set directly</span></span>

> [!NOTE]
> <span data-ttu-id="c78c9-115">Cette section s’applique uniquement aux API propres à l’hôte pour Excel et Word.</span><span class="sxs-lookup"><span data-stu-id="c78c9-115">This section only applies to the host-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="c78c9-116">Certaines propriétés ne peuvent pas être définies, bien qu’elles soient accessibles en écriture.</span><span class="sxs-lookup"><span data-stu-id="c78c9-116">Some properties cannot be set, despite being writable.</span></span> <span data-ttu-id="c78c9-117">Ces propriétés font partie d’une propriété parent qui doit être définie en tant qu’objet unique.</span><span class="sxs-lookup"><span data-stu-id="c78c9-117">These properties are part of a parent property that must be set as a single object.</span></span> <span data-ttu-id="c78c9-118">Cela est dû au fait que cette propriété Parent repose sur les sous-propriétés ayant des relations logiques spécifiques.</span><span class="sxs-lookup"><span data-stu-id="c78c9-118">This is because that parent property relies on the subproperties having specific, logical relationships.</span></span> <span data-ttu-id="c78c9-119">Ces propriétés parent doivent être définies à l’aide de la notation littérale d’objet pour définir l’objet entier, au lieu de définir les sous-propriétés individuelles de cet objet.</span><span class="sxs-lookup"><span data-stu-id="c78c9-119">These parent properties must be set using object literal notation to set the entire object, instead of setting that object's individual subproperties.</span></span> <span data-ttu-id="c78c9-120">Vous trouverez un exemple dans [PageLayout](/javascript/api/excel/excel.pagelayout).</span><span class="sxs-lookup"><span data-stu-id="c78c9-120">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="c78c9-121">La `zoom` propriété doit être définie avec un seul objet [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) , comme illustré ci-dessous :</span><span class="sxs-lookup"><span data-stu-id="c78c9-121">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="c78c9-122">Dans l’exemple précédent, vous ne seriez ***pas*** en mesure d' `zoom` affecter directement une `sheet.pageLayout.zoom.scale = 200;`valeur :.</span><span class="sxs-lookup"><span data-stu-id="c78c9-122">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="c78c9-123">Cette instruction génère une erreur car `zoom` elle n’est pas chargée.</span><span class="sxs-lookup"><span data-stu-id="c78c9-123">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="c78c9-124">Même si `zoom` elles ont été chargées, l’ensemble de l’étendue ne prendra pas effet.</span><span class="sxs-lookup"><span data-stu-id="c78c9-124">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="c78c9-125">Toutes les opérations de `zoom`contexte se produisent, actualisant l’objet proxy dans le complément et remplaçant les valeurs définies localement.</span><span class="sxs-lookup"><span data-stu-id="c78c9-125">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="c78c9-126">Ce comportement diffère des [Propriétés de navigation](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) telles que [Range. format](/javascript/api/excel/excel.range#format).</span><span class="sxs-lookup"><span data-stu-id="c78c9-126">This behavior differs from [navigational properties](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="c78c9-127">Les propriétés `format` de peuvent être définies à l’aide de la navigation d’objet, comme illustré ci-dessous :</span><span class="sxs-lookup"><span data-stu-id="c78c9-127">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="c78c9-128">Vous pouvez identifier une propriété qui ne peut pas avoir ses sous-propriétés directement définies en vérifiant son modificateur en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="c78c9-128">You can identify a property that cannot have its subproperties directly set by checking its read-only modifier.</span></span> <span data-ttu-id="c78c9-129">Les propriétés non en lecture seule de toutes les propriétés en lecture seule peuvent être définies directement.</span><span class="sxs-lookup"><span data-stu-id="c78c9-129">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="c78c9-130">Les propriétés accessibles en `PageLayout.zoom` écriture comme doivent être définies avec un objet à ce niveau.</span><span class="sxs-lookup"><span data-stu-id="c78c9-130">Writeable properties like `PageLayout.zoom` must be set with an object at that level.</span></span> <span data-ttu-id="c78c9-131">En Résumé :</span><span class="sxs-lookup"><span data-stu-id="c78c9-131">In summary:</span></span>

- <span data-ttu-id="c78c9-132">Propriété en lecture seule : les sous-propriétés peuvent être définies via la navigation.</span><span class="sxs-lookup"><span data-stu-id="c78c9-132">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="c78c9-133">Propriété accessible en écriture : les sous-propriétés ne peuvent pas être définies par le biais de la navigation (elles doivent être définies dans le cadre de l’attribution initiale de l’objet parent).</span><span class="sxs-lookup"><span data-stu-id="c78c9-133">Writable property: Subproperties cannot be set through navigation (must be set as part of the initial parent object assignment).</span></span>

## <a name="excel-data-transfer-limits"></a><span data-ttu-id="c78c9-134">Limites de transfert de données Excel</span><span class="sxs-lookup"><span data-stu-id="c78c9-134">Excel data transfer limits</span></span>

<span data-ttu-id="c78c9-135">Si vous créez un complément Excel, tenez compte des limitations de taille suivantes lors de l’interaction avec le classeur :</span><span class="sxs-lookup"><span data-stu-id="c78c9-135">If you're building an Excel add-in, be aware of the following size limitations when interacting with the workbook:</span></span>

- <span data-ttu-id="c78c9-136">Excel sur le web a une limite de taille de charge utile de 5 Mo pour les demandes et les réponses.</span><span class="sxs-lookup"><span data-stu-id="c78c9-136">Excel on the web has a payload size limit for requests and responses of 5MB.</span></span> <span data-ttu-id="c78c9-137">L’erreur `RichAPI.Error` est déclenchée en cas de dépassement de cette limite.</span><span class="sxs-lookup"><span data-stu-id="c78c9-137">`RichAPI.Error` will be thrown if that limit is exceeded.</span></span>
- <span data-ttu-id="c78c9-138">Une plage est limitée à 5 millions cellules pour les opérations Get.</span><span class="sxs-lookup"><span data-stu-id="c78c9-138">A range is limited to five million cells for get operations.</span></span>

<span data-ttu-id="c78c9-139">Si vous prévoyez que l’entrée de l’utilisateur dépasse ces limites, veillez à vérifier les `context.sync()`données avant d’appeler.</span><span class="sxs-lookup"><span data-stu-id="c78c9-139">If you expect user input to exceed these limits, be sure to check the data before calling `context.sync()`.</span></span> <span data-ttu-id="c78c9-140">Fractionnez l’opération en plusieurs parties si nécessaire.</span><span class="sxs-lookup"><span data-stu-id="c78c9-140">Split the operation into smaller pieces as needed.</span></span> <span data-ttu-id="c78c9-141">Veillez à appeler `context.sync()` pour chaque sous-opération afin d’éviter que ces opérations soient regroupées par lots.</span><span class="sxs-lookup"><span data-stu-id="c78c9-141">Be sure to call `context.sync()` for each sub-operation to avoid those operations getting batched together again.</span></span>

<span data-ttu-id="c78c9-142">Ces limitations sont généralement dépassées par les grandes plages.</span><span class="sxs-lookup"><span data-stu-id="c78c9-142">These limitations are typically exceeded by large ranges.</span></span> <span data-ttu-id="c78c9-143">Votre complément peut utiliser [RangeAreas](/javascript/api/excel/excel.rangeareas) pour mettre à jour les cellules dans une plage plus grande de manière stratégique.</span><span class="sxs-lookup"><span data-stu-id="c78c9-143">Your add-in might be able to use [RangeAreas](/javascript/api/excel/excel.rangeareas) to strategically update cells within a larger range.</span></span> <span data-ttu-id="c78c9-144">Pour plus d’informations, consultez [travailler simultanément avec plusieurs plages dans des compléments Excel](../excel/excel-add-ins-multiple-ranges.md) .</span><span class="sxs-lookup"><span data-stu-id="c78c9-144">See [Work with multiple ranges simultaneously in Excel add-ins](../excel/excel-add-ins-multiple-ranges.md) for more information.</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="c78c9-145">Définition de propriétés en lecture seule</span><span class="sxs-lookup"><span data-stu-id="c78c9-145">Setting read-only properties</span></span>

<span data-ttu-id="c78c9-146">Les [définitions](referencing-the-javascript-api-for-office-library-from-its-cdn.md) de la machine à écrire pour Office js spécifient les propriétés d’objet en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="c78c9-146">The [TypeScript definitions](referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="c78c9-147">Si vous tentez de définir une propriété en lecture seule, l’opération d’écriture échoue sans avertissement, sans qu’aucune erreur ne soit générée.</span><span class="sxs-lookup"><span data-stu-id="c78c9-147">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="c78c9-148">L’exemple suivant tente à tort de définir la propriété en lecture seule [Chart.ID](/javascript/api/excel/excel.chart#id).</span><span class="sxs-lookup"><span data-stu-id="c78c9-148">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="removing-event-handlers"></a><span data-ttu-id="c78c9-149">Suppression de gestionnaires d’événements</span><span class="sxs-lookup"><span data-stu-id="c78c9-149">Removing event handlers</span></span>

<span data-ttu-id="c78c9-150">Les gestionnaires d’événements doivent être supprimés à l' `RequestContext` aide du même que celui dans lequel ils ont été ajoutés.</span><span class="sxs-lookup"><span data-stu-id="c78c9-150">Event handlers must be removed using the same `RequestContext` in which they were added.</span></span> <span data-ttu-id="c78c9-151">Si vous avez besoin que votre complément supprime un gestionnaire d’événements en cours d’exécution, vous devez stocker l’objet Context utilisé pour ajouter le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="c78c9-151">If you need your add-in to remove an event handler while running, you'll need to store the context object used to add the handler.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="c78c9-152">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c78c9-152">See also</span></span>

- <span data-ttu-id="c78c9-153">[OfficeDev/Office-js](https://github.com/OfficeDev/office-js/issues): le lieu de signaler et d’afficher les problèmes liés à la plateforme des compléments Office et aux API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="c78c9-153">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="c78c9-154">[Débordement de pile](https://stackoverflow.com/questions/tagged/office-js): emplacement où poser des questions de programmation sur les API JavaScript Office.</span><span class="sxs-lookup"><span data-stu-id="c78c9-154">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="c78c9-155">Veillez à appliquer la balise « Office-js » à votre question lors de la publication dans le débordement de pile.</span><span class="sxs-lookup"><span data-stu-id="c78c9-155">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="c78c9-156">[UserVoice](https://officespdev.uservoice.com/): le lieu de suggérer de nouvelles fonctionnalités pour la plateforme des compléments Office et les API JavaScript pour Office.</span><span class="sxs-lookup"><span data-stu-id="c78c9-156">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
