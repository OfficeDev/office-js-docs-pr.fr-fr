---
title: Ensembles de conditions requises de l’API JavaScript pour OneNote
description: En savoir plus sur les ensembles de conditions requises de l’API JavaScript pour OneNote.
ms.date: 08/24/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: c8cadacac640cbe710c9894a65ee780267066afc
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293526"
---
# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="cb2ac-103">Ensembles de conditions requises de l’API JavaScript pour OneNote</span><span class="sxs-lookup"><span data-stu-id="cb2ac-103">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="cb2ac-p101">Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API qu’ils nécessitent. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="cb2ac-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="cb2ac-107">Le tableau suivant répertorie les ensembles de conditions requises pour OneNote, les applications clientes Office qui prennent en charge ces conditions ainsi que les numéros de version ou la date de disponibilité.</span><span class="sxs-lookup"><span data-stu-id="cb2ac-107">The following table lists the OneNote requirement sets, the Office client applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="cb2ac-108">Ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="cb2ac-108">Requirement set</span></span>  |  <span data-ttu-id="cb2ac-109">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="cb2ac-109">Office on the web</span></span> |
|:-----|:-----|
| [<span data-ttu-id="cb2ac-110">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="cb2ac-110">OneNoteApi 1.1</span></span>](/javascript/api/onenote?view=onenote-js-1.1)  | <span data-ttu-id="cb2ac-111">Septembre 2016</span><span class="sxs-lookup"><span data-stu-id="cb2ac-111">September 2016</span></span> |  

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="cb2ac-112">API JavaScript pour OneNote 1.1</span><span class="sxs-lookup"><span data-stu-id="cb2ac-112">OneNote JavaScript API 1.1</span></span>

<span data-ttu-id="cb2ac-113">L’API JavaScript 1.1 pour OneNote est la première version de l’API.</span><span class="sxs-lookup"><span data-stu-id="cb2ac-113">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="cb2ac-114">Pour plus d’informations sur l’API, consultez les rubriques de référence sur l’[Récapitulatif de programmation API JavaScript pour OneNote](../../onenote/onenote-add-ins-programming-overview.md).</span><span class="sxs-lookup"><span data-stu-id="cb2ac-114">For details about the API, see the [OneNote JavaScript API programming overview](../../onenote/onenote-add-ins-programming-overview.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="cb2ac-115">Vérification de la prise en charge d’une exigence d'exécution</span><span class="sxs-lookup"><span data-stu-id="cb2ac-115">Runtime requirement support check</span></span>

<span data-ttu-id="cb2ac-116">Lors de l’exécution, les compléments peuvent vérifier si une application Office spécifique prend en charge une série de conditions requises d’API en procédant comme suit.</span><span class="sxs-lookup"><span data-stu-id="cb2ac-116">At runtime, add-ins can check if a particular Office application supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="cb2ac-117">Vérification de la prise en charge des conditions requises basée sur le manifeste</span><span class="sxs-lookup"><span data-stu-id="cb2ac-117">Manifest-based requirement support check</span></span>

<span data-ttu-id="cb2ac-118">Utilisez l’élément `Requirements` dans le manifeste du complément pour spécifier des ensembles de conditions requises essentiels ou des membres d’API que votre complément doit utiliser.</span><span class="sxs-lookup"><span data-stu-id="cb2ac-118">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="cb2ac-119">Si l’application ou la plateforme Office ne prend pas en charge les ensembles de conditions requises ou les membres d’API spécifiés dans l’élément `Requirements`, le complément ne s’exécute pas dans cette application ou plateforme et ne s’affiche pas dans Mes compléments.</span><span class="sxs-lookup"><span data-stu-id="cb2ac-119">If the Office application or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that application or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="cb2ac-120">Cet exemple de code illustre un complément qui se charge dans toutes les applications clientes Office qui prennent en charge l’ensemble de conditions requises OneNoteApi, version 1.1.</span><span class="sxs-lookup"><span data-stu-id="cb2ac-120">The following code example shows an add-in that loads in all Office client applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="cb2ac-121">Séries de conditions requises des API communes pour Office</span><span class="sxs-lookup"><span data-stu-id="cb2ac-121">Office Common API requirement sets</span></span>

<span data-ttu-id="cb2ac-122">Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="cb2ac-122">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="cb2ac-123">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="cb2ac-123">See also</span></span>

- [<span data-ttu-id="cb2ac-124">Documentation référence de l’API JavaScript pour OneNote</span><span class="sxs-lookup"><span data-stu-id="cb2ac-124">OneNote JavaScript API reference documentation</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="cb2ac-125">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="cb2ac-125">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="cb2ac-126">Spécifier les exigences en matière d’applications Office et d’API</span><span class="sxs-lookup"><span data-stu-id="cb2ac-126">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="cb2ac-127">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="cb2ac-127">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
