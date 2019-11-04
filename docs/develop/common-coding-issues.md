---
title: Problèmes de codage courants et comportements de plateforme inattendus
description: Liste des problèmes de plateforme d’API JavaScript pour Office fréquemment rencontrés par les développeurs.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: d39c379961833cdb924628becf2c2da3f7e271b9
ms.sourcegitcommit: 59d29d01bce7543ebebf86e5a86db00cf54ca14a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/01/2019
ms.locfileid: "37924793"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="b51d0-103">Problèmes de codage courants et comportements de plateforme inattendus</span><span class="sxs-lookup"><span data-stu-id="b51d0-103">Common coding issues and unexpected platform behaviors</span></span>

<span data-ttu-id="b51d0-104">Cet article met en évidence les aspects de l’API JavaScript pour Office qui peuvent entraîner un comportement inattendu ou nécessiter des modèles de codage spécifiques pour obtenir le résultat souhaité.</span><span class="sxs-lookup"><span data-stu-id="b51d0-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="b51d0-105">Si vous rencontrez un problème qui se trouve dans cette liste, faites-le nous connaître en utilisant le formulaire de commentaires au bas de l’article.</span><span class="sxs-lookup"><span data-stu-id="b51d0-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a><span data-ttu-id="b51d0-106">Les API communes et les API Outlook ne sont pas basées sur la promesse</span><span class="sxs-lookup"><span data-stu-id="b51d0-106">Common APIs and Outlook APIs are not promise-based</span></span>

<span data-ttu-id="b51d0-107">Les [API communes](/javascript/api/office) (qui ne sont pas liées à un hôte Office particulier) et les [API Outlook](/javascript/api/outlook) utilisent un modèle de programmation basé sur les rappels.</span><span class="sxs-lookup"><span data-stu-id="b51d0-107">The [Common APIs](/javascript/api/office) (those that are not tied to a particular Office host) and [Outlook APIs](/javascript/api/outlook) use a callback-based programming model.</span></span> <span data-ttu-id="b51d0-108">L’interaction avec le document Office sous-jacent nécessite un appel asynchrone en lecture ou en écriture qui spécifie un rappel à exécuter lorsque l’opération se termine.</span><span class="sxs-lookup"><span data-stu-id="b51d0-108">Interacting with the underlying Office document requires an asynchronous read or write call that specifies a callback to be ran when the operation completes.</span></span> <span data-ttu-id="b51d0-109">Pour obtenir un exemple de ce modèle, consultez la rubrique [document. getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="b51d0-109">For an example of this pattern, see [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span>

<span data-ttu-id="b51d0-110">Ces méthodes d’API et d’API courantes ne renvoient pas de [promesses](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span><span class="sxs-lookup"><span data-stu-id="b51d0-110">These Common API and Outlook API methods do not return [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span></span> <span data-ttu-id="b51d0-111">Par conséquent, vous ne pouvez pas utiliser [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) pour suspendre l’exécution jusqu’à la fin de l’opération asynchrone.</span><span class="sxs-lookup"><span data-stu-id="b51d0-111">Therefore, you cannot use [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) to pause the execution until the asynchronous operation completes.</span></span> <span data-ttu-id="b51d0-112">Si vous avez `await` besoin de comportement, vous pouvez encapsuler l’appel de méthode dans une promesse créée de manière explicite.</span><span class="sxs-lookup"><span data-stu-id="b51d0-112">If you need `await` behavior, you can wrap the method call in an explicitly created Promise.</span></span>

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
> <span data-ttu-id="b51d0-113">La documentation de référence contient l’implémentation encapsulée de [fichier. getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span><span class="sxs-lookup"><span data-stu-id="b51d0-113">The reference documentation contains the Promise-wrapped implementation of [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span></span>

## <a name="some-properties-must-be-set-with-json-structs"></a><span data-ttu-id="b51d0-114">Certaines propriétés doivent être définies avec des structs JSON</span><span class="sxs-lookup"><span data-stu-id="b51d0-114">Some properties must be set with JSON structs</span></span>

> [!NOTE]
> <span data-ttu-id="b51d0-115">Cette section s’applique uniquement aux API propres à l’hôte pour Excel et Word.</span><span class="sxs-lookup"><span data-stu-id="b51d0-115">This section only applies to the host-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="b51d0-116">Certaines propriétés doivent être définies en tant que structs JSON, au lieu de définir leurs sous-propriétés individuelles.</span><span class="sxs-lookup"><span data-stu-id="b51d0-116">Some properties must be set as JSON structs, instead of setting their individual subproperties.</span></span> <span data-ttu-id="b51d0-117">Vous trouverez un exemple dans [PageLayout](/javascript/api/excel/excel.pagelayout).</span><span class="sxs-lookup"><span data-stu-id="b51d0-117">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="b51d0-118">La `zoom` propriété doit être définie avec un seul objet [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) , comme illustré ci-dessous :</span><span class="sxs-lookup"><span data-stu-id="b51d0-118">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="b51d0-119">Dans l’exemple précédent, vous ne seriez ***pas*** en mesure d' `zoom` affecter directement une `sheet.pageLayout.zoom.scale = 200;`valeur :.</span><span class="sxs-lookup"><span data-stu-id="b51d0-119">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="b51d0-120">Cette instruction génère une erreur car `zoom` elle n’est pas chargée.</span><span class="sxs-lookup"><span data-stu-id="b51d0-120">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="b51d0-121">Même si `zoom` elles ont été chargées, l’ensemble de l’étendue ne prendra pas effet.</span><span class="sxs-lookup"><span data-stu-id="b51d0-121">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="b51d0-122">Toutes les opérations de `zoom`contexte se produisent, actualisant l’objet proxy dans le complément et remplaçant les valeurs définies localement.</span><span class="sxs-lookup"><span data-stu-id="b51d0-122">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="b51d0-123">Ce comportement diffère des [Propriétés de navigation](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) telles que [Range. format](/javascript/api/excel/excel.range#format).</span><span class="sxs-lookup"><span data-stu-id="b51d0-123">This behavior differs from [navigational properties](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="b51d0-124">Les propriétés `format` de peuvent être définies à l’aide de la navigation d’objet, comme illustré ci-dessous :</span><span class="sxs-lookup"><span data-stu-id="b51d0-124">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="b51d0-125">Vous pouvez identifier une propriété dont les propriétés subordonnées doivent être définies avec un struct JSON en vérifiant son modificateur en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="b51d0-125">You can identify a property that must have its subproperties set with a JSON struct by checking its read-only modifier.</span></span> <span data-ttu-id="b51d0-126">Les propriétés non en lecture seule de toutes les propriétés en lecture seule peuvent être définies directement.</span><span class="sxs-lookup"><span data-stu-id="b51d0-126">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="b51d0-127">Les propriétés accessibles en `PageLayout.zoom` écriture comme doivent être définies avec une structure JSON.</span><span class="sxs-lookup"><span data-stu-id="b51d0-127">Writeable properties like `PageLayout.zoom` must be set with a JSON struct.</span></span> <span data-ttu-id="b51d0-128">En Résumé :</span><span class="sxs-lookup"><span data-stu-id="b51d0-128">In summary:</span></span>

- <span data-ttu-id="b51d0-129">Propriété en lecture seule : les sous-propriétés peuvent être définies via la navigation.</span><span class="sxs-lookup"><span data-stu-id="b51d0-129">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="b51d0-130">Propriété accessible en écriture : les sous-propriétés doivent être définies avec une structure JSON (et ne peuvent pas être définies via la navigation).</span><span class="sxs-lookup"><span data-stu-id="b51d0-130">Writable property: Subproperties must be set with a JSON struct (and cannot be set through navigation).</span></span>

## <a name="excel-range-limits"></a><span data-ttu-id="b51d0-131">Limites de plage Excel</span><span class="sxs-lookup"><span data-stu-id="b51d0-131">Excel Range limits</span></span>

<span data-ttu-id="b51d0-132">Si vous créez un complément Excel qui utilise des plages, gardez à l’esprit les limitations de taille suivantes :</span><span class="sxs-lookup"><span data-stu-id="b51d0-132">If you're building an Excel add-in that uses ranges, be aware of the following size limitations:</span></span>

- <span data-ttu-id="b51d0-133">Excel sur le web a une limite de taille de charge utile de 5 Mo pour les demandes et les réponses.</span><span class="sxs-lookup"><span data-stu-id="b51d0-133">Excel on the web has a payload size limit for requests and responses of 5MB.</span></span> <span data-ttu-id="b51d0-134">L’erreur `RichAPI.Error` est déclenchée en cas de dépassement de cette limite.</span><span class="sxs-lookup"><span data-stu-id="b51d0-134">`RichAPI.Error` will be thrown if that limit is exceeded.</span></span>
- <span data-ttu-id="b51d0-135">Une plage est limitée à 5 millions cellules pour les opérations Set.</span><span class="sxs-lookup"><span data-stu-id="b51d0-135">A range is limited to five million cells for set operations.</span></span>

<span data-ttu-id="b51d0-136">Si vous prévoyez que l’entrée de l’utilisateur dépasse ces limites, veillez à vérifier les données et à les fractionner en plusieurs objets.</span><span class="sxs-lookup"><span data-stu-id="b51d0-136">If you expect user input to exceed these limits, be sure to check the data and split the ranges into multiple objects.</span></span> <span data-ttu-id="b51d0-137">Vous devrez également envoyer plusieurs `context.sync()` appels afin d’éviter que les opérations de plage plus petites soient regroupées.</span><span class="sxs-lookup"><span data-stu-id="b51d0-137">You'll also need to submit multiple `context.sync()` calls to avoid the smaller range operations getting batched together again.</span></span>

<span data-ttu-id="b51d0-138">Votre complément peut utiliser [RangeAreas](/javascript/api/excel/excel.rangeareas) pour mettre à jour les cellules dans une plage plus grande de manière stratégique.</span><span class="sxs-lookup"><span data-stu-id="b51d0-138">Your add-in might be able to use [RangeAreas](/javascript/api/excel/excel.rangeareas) to strategically update cells within a larger range.</span></span> <span data-ttu-id="b51d0-139">Pour plus d’informations, consultez [travailler simultanément avec plusieurs plages dans des compléments Excel](../excel/excel-add-ins-multiple-ranges.md) .</span><span class="sxs-lookup"><span data-stu-id="b51d0-139">See [Work with multiple ranges simultaneously in Excel add-ins](../excel/excel-add-ins-multiple-ranges.md) for more information.</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="b51d0-140">Définition de propriétés en lecture seule</span><span class="sxs-lookup"><span data-stu-id="b51d0-140">Setting read-only properties</span></span>

<span data-ttu-id="b51d0-141">Les [définitions](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) de la machine à écrire pour Office js spécifient les propriétés d’objet en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="b51d0-141">The [TypeScript definitions](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="b51d0-142">Si vous tentez de définir une propriété en lecture seule, l’opération d’écriture échoue sans avertissement, sans qu’aucune erreur ne soit générée.</span><span class="sxs-lookup"><span data-stu-id="b51d0-142">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="b51d0-143">L’exemple suivant tente à tort de définir la propriété en lecture seule [Chart.ID](/javascript/api/excel/excel.chart#id).</span><span class="sxs-lookup"><span data-stu-id="b51d0-143">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="see-also"></a><span data-ttu-id="b51d0-144">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b51d0-144">See also</span></span>

- <span data-ttu-id="b51d0-145">[OfficeDev/Office-js](https://github.com/OfficeDev/office-js/issues): le lieu de signaler et d’afficher les problèmes liés à la plateforme des compléments Office et aux API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="b51d0-145">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="b51d0-146">[Débordement de pile](https://stackoverflow.com/questions/tagged/office-js): emplacement où poser des questions de programmation sur les API JavaScript Office.</span><span class="sxs-lookup"><span data-stu-id="b51d0-146">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="b51d0-147">Veillez à appliquer la balise « Office-js » à votre question lors de la publication dans le débordement de pile.</span><span class="sxs-lookup"><span data-stu-id="b51d0-147">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="b51d0-148">[UserVoice](https://officespdev.uservoice.com/): le lieu de suggérer de nouvelles fonctionnalités pour la plateforme des compléments Office et les API JavaScript pour Office.</span><span class="sxs-lookup"><span data-stu-id="b51d0-148">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
