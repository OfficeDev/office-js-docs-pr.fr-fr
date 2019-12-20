---
title: Problèmes de codage courants et comportements de plateforme inattendus
description: Liste des problèmes de plateforme d’API JavaScript pour Office fréquemment rencontrés par les développeurs.
ms.date: 12/05/2019
localization_priority: Normal
ms.openlocfilehash: 4271db2a9c61de419dd36fb0277574ffe0929c58
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814011"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="9ffba-103">Problèmes de codage courants et comportements de plateforme inattendus</span><span class="sxs-lookup"><span data-stu-id="9ffba-103">Common coding issues and unexpected platform behaviors</span></span>

<span data-ttu-id="9ffba-104">Cet article met en évidence les aspects de l’API JavaScript pour Office qui peuvent entraîner un comportement inattendu ou nécessiter des modèles de codage spécifiques pour obtenir le résultat souhaité.</span><span class="sxs-lookup"><span data-stu-id="9ffba-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="9ffba-105">Si vous rencontrez un problème qui se trouve dans cette liste, faites-le nous connaître en utilisant le formulaire de commentaires au bas de l’article.</span><span class="sxs-lookup"><span data-stu-id="9ffba-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a><span data-ttu-id="9ffba-106">Les API communes et les API Outlook ne sont pas basées sur la promesse</span><span class="sxs-lookup"><span data-stu-id="9ffba-106">Common APIs and Outlook APIs are not promise-based</span></span>

<span data-ttu-id="9ffba-107">Les [API communes](/javascript/api/office) (qui ne sont pas liées à un hôte Office particulier) et les [API Outlook](/javascript/api/outlook) utilisent un modèle de programmation basé sur les rappels.</span><span class="sxs-lookup"><span data-stu-id="9ffba-107">The [Common APIs](/javascript/api/office) (those that are not tied to a particular Office host) and [Outlook APIs](/javascript/api/outlook) use a callback-based programming model.</span></span> <span data-ttu-id="9ffba-108">L’interaction avec le document Office sous-jacent nécessite un appel asynchrone en lecture ou en écriture qui spécifie un rappel à exécuter lorsque l’opération se termine.</span><span class="sxs-lookup"><span data-stu-id="9ffba-108">Interacting with the underlying Office document requires an asynchronous read or write call that specifies a callback to be ran when the operation completes.</span></span> <span data-ttu-id="9ffba-109">Pour obtenir un exemple de ce modèle, consultez la rubrique [document. getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="9ffba-109">For an example of this pattern, see [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span>

<span data-ttu-id="9ffba-110">Ces méthodes d’API et d’API courantes ne renvoient pas de [promesses](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span><span class="sxs-lookup"><span data-stu-id="9ffba-110">These Common API and Outlook API methods do not return [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span></span> <span data-ttu-id="9ffba-111">Par conséquent, vous ne pouvez pas utiliser [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) pour suspendre l’exécution jusqu’à la fin de l’opération asynchrone.</span><span class="sxs-lookup"><span data-stu-id="9ffba-111">Therefore, you cannot use [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) to pause the execution until the asynchronous operation completes.</span></span> <span data-ttu-id="9ffba-112">Si vous avez `await` besoin de comportement, vous pouvez encapsuler l’appel de méthode dans une promesse créée de manière explicite.</span><span class="sxs-lookup"><span data-stu-id="9ffba-112">If you need `await` behavior, you can wrap the method call in an explicitly created Promise.</span></span>

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
> <span data-ttu-id="9ffba-113">La documentation de référence contient l’implémentation encapsulée de [fichier. getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span><span class="sxs-lookup"><span data-stu-id="9ffba-113">The reference documentation contains the Promise-wrapped implementation of [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span></span>

## <a name="some-properties-must-be-set-with-json-structs"></a><span data-ttu-id="9ffba-114">Certaines propriétés doivent être définies avec des structs JSON</span><span class="sxs-lookup"><span data-stu-id="9ffba-114">Some properties must be set with JSON structs</span></span>

> [!NOTE]
> <span data-ttu-id="9ffba-115">Cette section s’applique uniquement aux API propres à l’hôte pour Excel et Word.</span><span class="sxs-lookup"><span data-stu-id="9ffba-115">This section only applies to the host-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="9ffba-116">Certaines propriétés doivent être définies en tant que structs JSON, au lieu de définir leurs sous-propriétés individuelles.</span><span class="sxs-lookup"><span data-stu-id="9ffba-116">Some properties must be set as JSON structs, instead of setting their individual subproperties.</span></span> <span data-ttu-id="9ffba-117">Vous trouverez un exemple dans [PageLayout](/javascript/api/excel/excel.pagelayout).</span><span class="sxs-lookup"><span data-stu-id="9ffba-117">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="9ffba-118">La `zoom` propriété doit être définie avec un seul objet [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) , comme illustré ci-dessous :</span><span class="sxs-lookup"><span data-stu-id="9ffba-118">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="9ffba-119">Dans l’exemple précédent, vous ne seriez ***pas*** en mesure d' `zoom` affecter directement une `sheet.pageLayout.zoom.scale = 200;`valeur :.</span><span class="sxs-lookup"><span data-stu-id="9ffba-119">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="9ffba-120">Cette instruction génère une erreur car `zoom` elle n’est pas chargée.</span><span class="sxs-lookup"><span data-stu-id="9ffba-120">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="9ffba-121">Même si `zoom` elles ont été chargées, l’ensemble de l’étendue ne prendra pas effet.</span><span class="sxs-lookup"><span data-stu-id="9ffba-121">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="9ffba-122">Toutes les opérations de `zoom`contexte se produisent, actualisant l’objet proxy dans le complément et remplaçant les valeurs définies localement.</span><span class="sxs-lookup"><span data-stu-id="9ffba-122">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="9ffba-123">Ce comportement diffère des [Propriétés de navigation](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) telles que [Range. format](/javascript/api/excel/excel.range#format).</span><span class="sxs-lookup"><span data-stu-id="9ffba-123">This behavior differs from [navigational properties](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="9ffba-124">Les propriétés `format` de peuvent être définies à l’aide de la navigation d’objet, comme illustré ci-dessous :</span><span class="sxs-lookup"><span data-stu-id="9ffba-124">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="9ffba-125">Vous pouvez identifier une propriété dont les propriétés subordonnées doivent être définies avec un struct JSON en vérifiant son modificateur en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="9ffba-125">You can identify a property that must have its subproperties set with a JSON struct by checking its read-only modifier.</span></span> <span data-ttu-id="9ffba-126">Les propriétés non en lecture seule de toutes les propriétés en lecture seule peuvent être définies directement.</span><span class="sxs-lookup"><span data-stu-id="9ffba-126">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="9ffba-127">Les propriétés accessibles en `PageLayout.zoom` écriture comme doivent être définies avec une structure JSON.</span><span class="sxs-lookup"><span data-stu-id="9ffba-127">Writeable properties like `PageLayout.zoom` must be set with a JSON struct.</span></span> <span data-ttu-id="9ffba-128">En Résumé :</span><span class="sxs-lookup"><span data-stu-id="9ffba-128">In summary:</span></span>

- <span data-ttu-id="9ffba-129">Propriété en lecture seule : les sous-propriétés peuvent être définies via la navigation.</span><span class="sxs-lookup"><span data-stu-id="9ffba-129">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="9ffba-130">Propriété accessible en écriture : les sous-propriétés doivent être définies avec une structure JSON (et ne peuvent pas être définies via la navigation).</span><span class="sxs-lookup"><span data-stu-id="9ffba-130">Writable property: Subproperties must be set with a JSON struct (and cannot be set through navigation).</span></span>

## <a name="excel-data-transfer-limits"></a><span data-ttu-id="9ffba-131">Limites de transfert de données Excel</span><span class="sxs-lookup"><span data-stu-id="9ffba-131">Excel data transfer limits</span></span>

<span data-ttu-id="9ffba-132">Si vous créez un complément Excel, tenez compte des limitations de taille suivantes lors de l’interaction avec le classeur :</span><span class="sxs-lookup"><span data-stu-id="9ffba-132">If you're building an Excel add-in, be aware of the following size limitations when interacting with the workbook:</span></span>

- <span data-ttu-id="9ffba-133">Excel sur le web a une limite de taille de charge utile de 5 Mo pour les demandes et les réponses.</span><span class="sxs-lookup"><span data-stu-id="9ffba-133">Excel on the web has a payload size limit for requests and responses of 5MB.</span></span> <span data-ttu-id="9ffba-134">L’erreur `RichAPI.Error` est déclenchée en cas de dépassement de cette limite.</span><span class="sxs-lookup"><span data-stu-id="9ffba-134">`RichAPI.Error` will be thrown if that limit is exceeded.</span></span>
- <span data-ttu-id="9ffba-135">Une plage est limitée à 5 millions cellules pour les opérations Get.</span><span class="sxs-lookup"><span data-stu-id="9ffba-135">A range is limited to five million cells for get operations.</span></span>

<span data-ttu-id="9ffba-136">Si vous prévoyez que l’entrée de l’utilisateur dépasse ces limites, veillez à vérifier les `context.sync()`données avant d’appeler.</span><span class="sxs-lookup"><span data-stu-id="9ffba-136">If you expect user input to exceed these limits, be sure to check the data before calling `context.sync()`.</span></span> <span data-ttu-id="9ffba-137">Fractionnez l’opération en plusieurs parties si nécessaire.</span><span class="sxs-lookup"><span data-stu-id="9ffba-137">Split the operation into smaller pieces as needed.</span></span> <span data-ttu-id="9ffba-138">Veillez à appeler `context.sync()` pour chaque sous-opération afin d’éviter que ces opérations soient regroupées par lots.</span><span class="sxs-lookup"><span data-stu-id="9ffba-138">Be sure to call `context.sync()` for each sub-operation to avoid those operations getting batched together again.</span></span>

<span data-ttu-id="9ffba-139">Ces limitations sont généralement dépassées par les grandes plages.</span><span class="sxs-lookup"><span data-stu-id="9ffba-139">These limitations are typically exceeded by large ranges.</span></span> <span data-ttu-id="9ffba-140">Votre complément peut utiliser [RangeAreas](/javascript/api/excel/excel.rangeareas) pour mettre à jour les cellules dans une plage plus grande de manière stratégique.</span><span class="sxs-lookup"><span data-stu-id="9ffba-140">Your add-in might be able to use [RangeAreas](/javascript/api/excel/excel.rangeareas) to strategically update cells within a larger range.</span></span> <span data-ttu-id="9ffba-141">Pour plus d’informations, consultez [travailler simultanément avec plusieurs plages dans des compléments Excel](../excel/excel-add-ins-multiple-ranges.md) .</span><span class="sxs-lookup"><span data-stu-id="9ffba-141">See [Work with multiple ranges simultaneously in Excel add-ins](../excel/excel-add-ins-multiple-ranges.md) for more information.</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="9ffba-142">Définition de propriétés en lecture seule</span><span class="sxs-lookup"><span data-stu-id="9ffba-142">Setting read-only properties</span></span>

<span data-ttu-id="9ffba-143">Les [définitions](referencing-the-javascript-api-for-office-library-from-its-cdn.md) de la machine à écrire pour Office js spécifient les propriétés d’objet en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="9ffba-143">The [TypeScript definitions](referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="9ffba-144">Si vous tentez de définir une propriété en lecture seule, l’opération d’écriture échoue sans avertissement, sans qu’aucune erreur ne soit générée.</span><span class="sxs-lookup"><span data-stu-id="9ffba-144">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="9ffba-145">L’exemple suivant tente à tort de définir la propriété en lecture seule [Chart.ID](/javascript/api/excel/excel.chart#id).</span><span class="sxs-lookup"><span data-stu-id="9ffba-145">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="removing-event-handlers"></a><span data-ttu-id="9ffba-146">Suppression de gestionnaires d’événements</span><span class="sxs-lookup"><span data-stu-id="9ffba-146">Removing event handlers</span></span>

<span data-ttu-id="9ffba-147">Les gestionnaires d’événements doivent être supprimés à l' `RequestContext` aide du même que celui dans lequel ils ont été ajoutés.</span><span class="sxs-lookup"><span data-stu-id="9ffba-147">Event handlers must be removed using the same `RequestContext` in which they were added.</span></span> <span data-ttu-id="9ffba-148">Si vous avez besoin que votre complément supprime un gestionnaire d’événements en cours d’exécution, vous devez stocker l’objet Context utilisé pour ajouter le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="9ffba-148">If you need your add-in to remove an event handler while running, you'll need to store the context object used to add the handler.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="9ffba-149">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="9ffba-149">See also</span></span>

- <span data-ttu-id="9ffba-150">[OfficeDev/Office-js](https://github.com/OfficeDev/office-js/issues): le lieu de signaler et d’afficher les problèmes liés à la plateforme des compléments Office et aux API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="9ffba-150">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="9ffba-151">[Débordement de pile](https://stackoverflow.com/questions/tagged/office-js): emplacement où poser des questions de programmation sur les API JavaScript Office.</span><span class="sxs-lookup"><span data-stu-id="9ffba-151">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="9ffba-152">Veillez à appliquer la balise « Office-js » à votre question lors de la publication dans le débordement de pile.</span><span class="sxs-lookup"><span data-stu-id="9ffba-152">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="9ffba-153">[UserVoice](https://officespdev.uservoice.com/): le lieu de suggérer de nouvelles fonctionnalités pour la plateforme des compléments Office et les API JavaScript pour Office.</span><span class="sxs-lookup"><span data-stu-id="9ffba-153">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
