---
title: Séries de conditions requises de l’API JavaScript pour PowerPoint
description: En savoir plus sur les ensembles de conditions requises de l’API JavaScript pour PowerPoint
ms.date: 03/11/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: 8349ff4c96e73ce93a77c86b709081f9355b1f06
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094301"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="b60d8-103">Séries de conditions requises de l’API JavaScript pour PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b60d8-103">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="b60d8-104">Requirement sets are named groups of API members.</span><span class="sxs-lookup"><span data-stu-id="b60d8-104">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="b60d8-105">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span><span class="sxs-lookup"><span data-stu-id="b60d8-105">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="b60d8-106">For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="b60d8-106">For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="b60d8-107">Le tableau suivant répertorie les séries de conditions requises pour PowerPoint, les applications hôtes Office qui prennent en charge ces conditions et les numéros de version ou la date de disponibilité.</span><span class="sxs-lookup"><span data-stu-id="b60d8-107">The following table lists the PowerPoint requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="b60d8-108">Ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="b60d8-108">Requirement set</span></span>  |  <span data-ttu-id="b60d8-109">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="b60d8-109">Office on Windows</span></span><br><span data-ttu-id="b60d8-110">(connecté à l’abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="b60d8-110">(connected to Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="b60d8-111">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="b60d8-111">Office on iPad</span></span><br><span data-ttu-id="b60d8-112">(connecté à l’abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="b60d8-112">(connected to Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="b60d8-113">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="b60d8-113">Office on Mac</span></span><br><span data-ttu-id="b60d8-114">(connecté à l’abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="b60d8-114">(connected to Microsoft 365 subscription)</span></span>  | <span data-ttu-id="b60d8-115">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="b60d8-115">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="b60d8-116">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="b60d8-116">PowerPointApi 1.1</span></span> | <span data-ttu-id="b60d8-117">Version 1810 (Build 11001.20074) ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="b60d8-117">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="b60d8-118">2.17 ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="b60d8-118">2.17 or later</span></span> | <span data-ttu-id="b60d8-119">16.19 ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="b60d8-119">16.19 or later</span></span> | <span data-ttu-id="b60d8-120">Octobre 2018</span><span class="sxs-lookup"><span data-stu-id="b60d8-120">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="b60d8-121">Numéros de version et de build d’Office</span><span class="sxs-lookup"><span data-stu-id="b60d8-121">Office versions and build numbers</span></span>

<span data-ttu-id="b60d8-122">Pour plus d’informations sur les versions et les numéros de build d’Office, voir :</span><span class="sxs-lookup"><span data-stu-id="b60d8-122">For more information about Office versions and build numbers, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="b60d8-123">API JavaScript pour PowerPoint 1.1</span><span class="sxs-lookup"><span data-stu-id="b60d8-123">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="b60d8-124">L’API JavaScript PowerPoint 1.1 inclut une seule API pour créer une nouvelle présentation.</span><span class="sxs-lookup"><span data-stu-id="b60d8-124">PowerPoint JavaScript API 1.1 contains a single API to create a new presentation.</span></span> <span data-ttu-id="b60d8-125">Pour plus de détails sur l’API, voir [API JavaScript pour PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="b60d8-125">For details about the API, see [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="b60d8-126">Vérification de la prise en charge d’une exigence d'exécution</span><span class="sxs-lookup"><span data-stu-id="b60d8-126">Runtime requirement support check</span></span>

<span data-ttu-id="b60d8-127">Lors de l’exécution, les compléments peuvent vérifier si un hôte particulier prend en charge une série de conditions requises d’API en procédant comme suit.</span><span class="sxs-lookup"><span data-stu-id="b60d8-127">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="b60d8-128">Vérification de la prise en charge des conditions requises basée sur le manifeste</span><span class="sxs-lookup"><span data-stu-id="b60d8-128">Manifest-based requirement support check</span></span>

<span data-ttu-id="b60d8-129">Utilisez l’élément `Requirements` dans le manifeste du complément pour spécifier des ensembles de conditions requises essentiels ou des membres d’API que votre complément doit utiliser.</span><span class="sxs-lookup"><span data-stu-id="b60d8-129">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="b60d8-130">Si la plateforme ou l’hôte Office ne prend pas en charge les ensembles de conditions requises ou les membres d’API spécifiés dans l’élément `Requirements`, le complément ne s’exécute pas dans cet hôte ou cette plateforme et ne s’affiche pas dans Mes compléments.</span><span class="sxs-lookup"><span data-stu-id="b60d8-130">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="b60d8-131">Cet exemple de code illustre un complément qui se charge dans toutes les applications hôtes Office qui prennent en charge l’ensemble de conditions requises OneNoteApi, version 1.1.</span><span class="sxs-lookup"><span data-stu-id="b60d8-131">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="b60d8-132">Séries de conditions requises des API communes pour Office</span><span class="sxs-lookup"><span data-stu-id="b60d8-132">Office Common API requirement sets</span></span>

<span data-ttu-id="b60d8-133">La plupart des fonctionnalités du complément PowerPoint proviennent de la série courante d’API.</span><span class="sxs-lookup"><span data-stu-id="b60d8-133">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="b60d8-134">Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="b60d8-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="b60d8-135">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b60d8-135">See also</span></span>

- [<span data-ttu-id="b60d8-136">Documentation référence de l’API JavaScript pour PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b60d8-136">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="b60d8-137">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="b60d8-137">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="b60d8-138">Spécification des exigences en matière d’hôtes Office et d’API</span><span class="sxs-lookup"><span data-stu-id="b60d8-138">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="b60d8-139">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="b60d8-139">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
