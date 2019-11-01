---
title: Problèmes de codage courants et comportements de plateforme inattendus
description: Liste des problèmes de plateforme d’API JavaScript pour Office fréquemment rencontrés par les développeurs.
ms.date: 10/29/2019
localization_priority: Normal
ms.openlocfilehash: 8cea95e3214585ba8e0b77535916f9c564dde9df
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902152"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="ec385-103">Problèmes de codage courants et comportements de plateforme inattendus</span><span class="sxs-lookup"><span data-stu-id="ec385-103">Common coding issues and unexpected platform behaviors</span></span>

<span data-ttu-id="ec385-104">Cet article met en évidence les aspects de l’API JavaScript pour Office qui peuvent entraîner un comportement inattendu ou nécessiter des modèles de codage spécifiques pour obtenir le résultat souhaité.</span><span class="sxs-lookup"><span data-stu-id="ec385-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="ec385-105">Si vous rencontrez un problème qui se trouve dans cette liste, faites-le nous connaître en utilisant le formulaire de commentaires au bas de l’article.</span><span class="sxs-lookup"><span data-stu-id="ec385-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="some-properties-must-be-set-with-json-structs"></a><span data-ttu-id="ec385-106">Certaines propriétés doivent être définies avec des structs JSON</span><span class="sxs-lookup"><span data-stu-id="ec385-106">Some properties must be set with JSON structs</span></span>

> [!NOTE]
> <span data-ttu-id="ec385-107">Cette section s’applique uniquement aux API propres à l’hôte pour Excel et Word.</span><span class="sxs-lookup"><span data-stu-id="ec385-107">This section only applies to the host-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="ec385-108">Certaines propriétés doivent être définies en tant que structs JSON, au lieu de définir leurs sous-propriétés individuelles.</span><span class="sxs-lookup"><span data-stu-id="ec385-108">Some properties must be set as JSON structs, instead of setting their individual subproperties.</span></span> <span data-ttu-id="ec385-109">Vous trouverez un exemple dans [PageLayout](/javascript/api/excel/excel.pagelayout).</span><span class="sxs-lookup"><span data-stu-id="ec385-109">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="ec385-110">La `zoom` propriété doit être définie avec un seul objet [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) , comme illustré ci-dessous :</span><span class="sxs-lookup"><span data-stu-id="ec385-110">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="ec385-111">Dans l’exemple précédent, vous ne seriez ***pas*** en mesure d' `zoom` affecter directement une `sheet.pageLayout.zoom.scale = 200;`valeur :.</span><span class="sxs-lookup"><span data-stu-id="ec385-111">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="ec385-112">Cette instruction génère une erreur car `zoom` elle n’est pas chargée.</span><span class="sxs-lookup"><span data-stu-id="ec385-112">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="ec385-113">Même si `zoom` elles ont été chargées, l’ensemble de l’étendue ne prendra pas effet.</span><span class="sxs-lookup"><span data-stu-id="ec385-113">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="ec385-114">Toutes les opérations de `zoom`contexte se produisent, actualisant l’objet proxy dans le complément et remplaçant les valeurs définies localement.</span><span class="sxs-lookup"><span data-stu-id="ec385-114">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="ec385-115">Ce comportement diffère des [Propriétés de navigation](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) telles que [Range. format](/javascript/api/excel/excel.range#format).</span><span class="sxs-lookup"><span data-stu-id="ec385-115">This behavior differs from [navigational properties](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="ec385-116">Les propriétés `format` de peuvent être définies à l’aide de la navigation d’objet, comme illustré ci-dessous :</span><span class="sxs-lookup"><span data-stu-id="ec385-116">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="ec385-117">Vous pouvez identifier une propriété dont les propriétés subordonnées doivent être définies avec un struct JSON en vérifiant son modificateur en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="ec385-117">You can identify a property that must have its subproperties set with a JSON struct by checking its read-only modifier.</span></span> <span data-ttu-id="ec385-118">Les propriétés non en lecture seule de toutes les propriétés en lecture seule peuvent être définies directement.</span><span class="sxs-lookup"><span data-stu-id="ec385-118">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="ec385-119">Les propriétés accessibles en `PageLayout.zoom` écriture comme doivent être définies avec une structure JSON.</span><span class="sxs-lookup"><span data-stu-id="ec385-119">Writeable properties like `PageLayout.zoom` must be set with a JSON struct.</span></span> <span data-ttu-id="ec385-120">En Résumé :</span><span class="sxs-lookup"><span data-stu-id="ec385-120">In summary:</span></span>

- <span data-ttu-id="ec385-121">Propriété en lecture seule : les sous-propriétés peuvent être définies via la navigation.</span><span class="sxs-lookup"><span data-stu-id="ec385-121">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="ec385-122">Propriété accessible en écriture : les sous-propriétés doivent être définies avec une structure JSON (et ne peuvent pas être définies via la navigation).</span><span class="sxs-lookup"><span data-stu-id="ec385-122">Writable property: Subproperties must be set with a JSON struct (and cannot be set through navigation).</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="ec385-123">Définition de propriétés en lecture seule</span><span class="sxs-lookup"><span data-stu-id="ec385-123">Setting read-only properties</span></span>

<span data-ttu-id="ec385-124">Les [définitions](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) de la machine à écrire pour Office js spécifient les propriétés d’objet en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="ec385-124">The [TypeScript definitions](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="ec385-125">Si vous tentez de définir une propriété en lecture seule, l’opération d’écriture échoue sans avertissement, sans qu’aucune erreur ne soit générée.</span><span class="sxs-lookup"><span data-stu-id="ec385-125">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="ec385-126">L’exemple suivant tente à tort de définir la propriété en lecture seule [Chart.ID](/javascript/api/excel/excel.chart#id).</span><span class="sxs-lookup"><span data-stu-id="ec385-126">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="see-also"></a><span data-ttu-id="ec385-127">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ec385-127">See also</span></span>

- <span data-ttu-id="ec385-128">[OfficeDev/Office-js](https://github.com/OfficeDev/office-js/issues): le lieu de signaler et d’afficher les problèmes liés à la plateforme des compléments Office et aux API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="ec385-128">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="ec385-129">[Débordement de pile](https://stackoverflow.com/questions/tagged/office-js): emplacement où poser des questions de programmation sur les API JavaScript Office.</span><span class="sxs-lookup"><span data-stu-id="ec385-129">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="ec385-130">Veillez à appliquer la balise « Office-js » à votre question lors de la publication dans le débordement de pile.</span><span class="sxs-lookup"><span data-stu-id="ec385-130">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="ec385-131">[UserVoice](https://officespdev.uservoice.com/): le lieu de suggérer de nouvelles fonctionnalités pour la plateforme des compléments Office et les API JavaScript pour Office.</span><span class="sxs-lookup"><span data-stu-id="ec385-131">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
