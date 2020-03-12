---
title: Ensembles de conditions requises de l’API JavaScript pour OneNote
description: ''
ms.date: 07/17/2019
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: d936d5f0c7c40cf79442eac76dbb9d94748a37a8
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596948"
---
# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="f0ce6-102">Ensembles de conditions requises de l’API JavaScript pour OneNote</span><span class="sxs-lookup"><span data-stu-id="f0ce6-102">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="f0ce6-p101">Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="f0ce6-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="f0ce6-106">Le tableau suivant répertorie les ensembles de conditions requises pour OneNote, les applications hôtes Office qui prennent en charge ces conditions et les numéros de version ou la date de disponibilité.</span><span class="sxs-lookup"><span data-stu-id="f0ce6-106">The following table lists the OneNote requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="f0ce6-107">Ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="f0ce6-107">Requirement set</span></span>  |  <span data-ttu-id="f0ce6-108">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="f0ce6-108">Office on the web</span></span> |
|:-----|:-----|
| [<span data-ttu-id="f0ce6-109">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f0ce6-109">OneNoteApi 1.1</span></span>](/javascript/api/onenote?view=onenote-js-1.1)  | <span data-ttu-id="f0ce6-110">Septembre 2016</span><span class="sxs-lookup"><span data-stu-id="f0ce6-110">September 2016</span></span> |  

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="f0ce6-111">Ensembles de conditions requises des API communes pour Office</span><span class="sxs-lookup"><span data-stu-id="f0ce6-111">Office Common API requirement sets</span></span>

<span data-ttu-id="f0ce6-112">Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="f0ce6-112">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="f0ce6-113">API JavaScript pour OneNote 1.1</span><span class="sxs-lookup"><span data-stu-id="f0ce6-113">OneNote JavaScript API 1.1</span></span>

<span data-ttu-id="f0ce6-114">L’API JavaScript 1.1 pour OneNote est la première version de l’API.</span><span class="sxs-lookup"><span data-stu-id="f0ce6-114">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="f0ce6-115">Pour plus d’informations sur l’API, consultez les rubriques de référence sur l’[Récapitulatif de programmation API JavaScript pour OneNote](../../onenote/onenote-add-ins-programming-overview.md).</span><span class="sxs-lookup"><span data-stu-id="f0ce6-115">For details about the API, see the [OneNote JavaScript API programming overview](../../onenote/onenote-add-ins-programming-overview.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="f0ce6-116">Vérification de la prise en charge d’un ensemble de conditions requises à l’exécution</span><span class="sxs-lookup"><span data-stu-id="f0ce6-116">Runtime requirement support check</span></span>

<span data-ttu-id="f0ce6-117">Lors de l’exécution, les compléments peuvent vérifier si un hôte particulier prend en charge une série de conditions requises d’API en procédant comme suit.</span><span class="sxs-lookup"><span data-stu-id="f0ce6-117">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="f0ce6-118">Vérification de la prise en charge des conditions requises basée sur le manifeste</span><span class="sxs-lookup"><span data-stu-id="f0ce6-118">Manifest-based requirement support check</span></span>

<span data-ttu-id="f0ce6-119">Utilisez l’élément `Requirements` dans le manifeste du complément pour spécifier des ensembles de conditions requises essentiels ou des membres d’API que votre complément doit utiliser.</span><span class="sxs-lookup"><span data-stu-id="f0ce6-119">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="f0ce6-120">Si la plateforme ou l’hôte Office ne prend pas en charge les ensembles de conditions requises ou les membres d’API spécifiés dans l’élément `Requirements`, le complément ne s’exécute pas dans cet hôte ou cette plateforme et ne s’affiche pas dans Mes compléments.</span><span class="sxs-lookup"><span data-stu-id="f0ce6-120">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="f0ce6-121">Cet exemple de code illustre un complément qui se charge dans toutes les applications hôtes Office qui prennent en charge l’ensemble de conditions requises OneNoteApi, version 1.1.</span><span class="sxs-lookup"><span data-stu-id="f0ce6-121">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="f0ce6-122">Ensembles de conditions requises des API communes pour Office</span><span class="sxs-lookup"><span data-stu-id="f0ce6-122">Office Common API requirement sets</span></span>

<span data-ttu-id="f0ce6-123">Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="f0ce6-123">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="f0ce6-124">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f0ce6-124">See also</span></span>

- [<span data-ttu-id="f0ce6-125">Documentation référence de l’API JavaScript pour OneNote</span><span class="sxs-lookup"><span data-stu-id="f0ce6-125">OneNote JavaScript API reference documentation</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="f0ce6-126">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="f0ce6-126">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="f0ce6-127">Spécification des exigences en matière d’hôtes Office et d’API</span><span class="sxs-lookup"><span data-stu-id="f0ce6-127">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="f0ce6-128">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="f0ce6-128">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
