---
title: Séries de conditions requises de l’API JavaScript pour PowerPoint
description: ''
ms.date: 03/11/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: ef76077c3a2a975fae8a0dc101e8e1b42ef66094
ms.sourcegitcommit: 05b73cdec5f4db7f0b8d48a5a552ee296a0332ca
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42600696"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="10399-102">Séries de conditions requises de l’API JavaScript pour PowerPoint</span><span class="sxs-lookup"><span data-stu-id="10399-102">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="10399-p101">Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="10399-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="10399-106">Le tableau suivant répertorie les séries de conditions requises pour PowerPoint, les applications hôtes Office qui prennent en charge ces conditions et les numéros de version ou la date de disponibilité.</span><span class="sxs-lookup"><span data-stu-id="10399-106">The following table lists the PowerPoint requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="10399-107">Ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="10399-107">Requirement set</span></span>  |  <span data-ttu-id="10399-108">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="10399-108">Office on Windows</span></span><br><span data-ttu-id="10399-109">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="10399-109">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="10399-110">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="10399-110">Office on iPad</span></span><br><span data-ttu-id="10399-111">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="10399-111">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="10399-112">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="10399-112">Office on Mac</span></span><br><span data-ttu-id="10399-113">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="10399-113">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="10399-114">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="10399-114">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="10399-115">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="10399-115">PowerPointApi 1.1</span></span> | <span data-ttu-id="10399-116">Version 1810 (Build 11001.20074) ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="10399-116">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="10399-117">2.17 ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="10399-117">2.17 or later</span></span> | <span data-ttu-id="10399-118">16.19 ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="10399-118">16.19 or later</span></span> | <span data-ttu-id="10399-119">Octobre 2018</span><span class="sxs-lookup"><span data-stu-id="10399-119">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="10399-120">Numéros de version et de build d’Office</span><span class="sxs-lookup"><span data-stu-id="10399-120">Office versions and build numbers</span></span>

<span data-ttu-id="10399-121">Pour plus d’informations sur les versions et les numéros de build d’Office, voir :</span><span class="sxs-lookup"><span data-stu-id="10399-121">For more information about Office versions and build numbers, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="10399-122">API JavaScript pour PowerPoint 1.1</span><span class="sxs-lookup"><span data-stu-id="10399-122">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="10399-123">L’API JavaScript PowerPoint 1.1 inclut une seule API pour créer une nouvelle présentation.</span><span class="sxs-lookup"><span data-stu-id="10399-123">PowerPoint JavaScript API 1.1 contains a single API to create a new presentation.</span></span> <span data-ttu-id="10399-124">Pour plus de détails sur l’API, voir [API JavaScript pour PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="10399-124">For details about the API, see [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="10399-125">Vérification de la prise en charge d’une exigence d'exécution</span><span class="sxs-lookup"><span data-stu-id="10399-125">Runtime requirement support check</span></span>

<span data-ttu-id="10399-126">Lors de l’exécution, les compléments peuvent vérifier si un hôte particulier prend en charge une série de conditions requises d’API en procédant comme suit.</span><span class="sxs-lookup"><span data-stu-id="10399-126">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="10399-127">Vérification de la prise en charge des conditions requises basée sur le manifeste</span><span class="sxs-lookup"><span data-stu-id="10399-127">Manifest-based requirement support check</span></span>

<span data-ttu-id="10399-128">Utilisez l’élément `Requirements` dans le manifeste du complément pour spécifier des ensembles de conditions requises essentiels ou des membres d’API que votre complément doit utiliser.</span><span class="sxs-lookup"><span data-stu-id="10399-128">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="10399-129">Si la plateforme ou l’hôte Office ne prend pas en charge les ensembles de conditions requises ou les membres d’API spécifiés dans l’élément `Requirements`, le complément ne s’exécute pas dans cet hôte ou cette plateforme et ne s’affiche pas dans Mes compléments.</span><span class="sxs-lookup"><span data-stu-id="10399-129">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="10399-130">Cet exemple de code illustre un complément qui se charge dans toutes les applications hôtes Office qui prennent en charge l’ensemble de conditions requises OneNoteApi, version 1.1.</span><span class="sxs-lookup"><span data-stu-id="10399-130">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="10399-131">Séries de conditions requises des API communes pour Office</span><span class="sxs-lookup"><span data-stu-id="10399-131">Office Common API requirement sets</span></span>

<span data-ttu-id="10399-132">La plupart des fonctionnalités du complément PowerPoint proviennent de la série courante d’API.</span><span class="sxs-lookup"><span data-stu-id="10399-132">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="10399-133">Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="10399-133">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="10399-134">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="10399-134">See also</span></span>

- [<span data-ttu-id="10399-135">Documentation référence de l’API JavaScript pour PowerPoint</span><span class="sxs-lookup"><span data-stu-id="10399-135">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="10399-136">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="10399-136">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="10399-137">Spécification des exigences en matière d’hôtes Office et d’API</span><span class="sxs-lookup"><span data-stu-id="10399-137">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="10399-138">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="10399-138">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
