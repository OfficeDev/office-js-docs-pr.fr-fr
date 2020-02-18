---
title: Séries de conditions requises de l’API JavaScript pour PowerPoint
description: ''
ms.date: 07/26/2019
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: 5bba2354cabba3c3ccd4ddf38d3e03c25a32b8a9
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950956"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="a750e-102">Séries de conditions requises de l’API JavaScript pour PowerPoint</span><span class="sxs-lookup"><span data-stu-id="a750e-102">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="a750e-p101">Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="a750e-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="a750e-106">Le tableau suivant répertorie les séries de conditions requises pour PowerPoint, les applications hôtes Office qui prennent en charge ces conditions et les numéros de version ou la date de disponibilité.</span><span class="sxs-lookup"><span data-stu-id="a750e-106">The following table lists the PowerPoint requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="a750e-107">Ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="a750e-107">Requirement set</span></span>  |  <span data-ttu-id="a750e-108">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="a750e-108">Office on Windows</span></span><br><span data-ttu-id="a750e-109">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="a750e-109">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="a750e-110">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="a750e-110">Office on iPad</span></span><br><span data-ttu-id="a750e-111">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="a750e-111">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="a750e-112">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="a750e-112">Office on Mac</span></span><br><span data-ttu-id="a750e-113">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="a750e-113">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="a750e-114">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="a750e-114">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="a750e-115">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="a750e-115">PowerPointApi 1.1</span></span> | <span data-ttu-id="a750e-116">Version 1810 (Build 11001.20074) ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="a750e-116">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="a750e-117">2.17 ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="a750e-117">2.17 or later</span></span> | <span data-ttu-id="a750e-118">16.19 ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="a750e-118">16.19 or later</span></span> | <span data-ttu-id="a750e-119">Octobre 2018</span><span class="sxs-lookup"><span data-stu-id="a750e-119">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="a750e-120">Numéros de version et de build d’Office</span><span class="sxs-lookup"><span data-stu-id="a750e-120">Office versions and build numbers</span></span>

<span data-ttu-id="a750e-121">Pour plus d’informations sur les versions et les numéros de build d’Office, voir :</span><span class="sxs-lookup"><span data-stu-id="a750e-121">For more information about Office versions and build numbers, see:</span></span>

- [<span data-ttu-id="a750e-122">Numéros de version et de build des canaux de réception des mises à jour pour les clients Office 365</span><span class="sxs-lookup"><span data-stu-id="a750e-122">Version and build numbers of update channel releases for Office 365 clients</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="a750e-123">Quelle est la version d’Office que j’utilise ?</span><span class="sxs-lookup"><span data-stu-id="a750e-123">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="a750e-124">Où trouver le numéro de version et de build pour une application cliente Office 365</span><span class="sxs-lookup"><span data-stu-id="a750e-124">Where you can find the version and build number for an Office 365 client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="a750e-125">API JavaScript pour PowerPoint 1.1</span><span class="sxs-lookup"><span data-stu-id="a750e-125">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="a750e-126">L’API JavaScript PowerPoint 1.1 inclut une seule API pour créer une nouvelle présentation.</span><span class="sxs-lookup"><span data-stu-id="a750e-126">PowerPoint JavaScript API 1.1 contains a single API to create a new presentation.</span></span> <span data-ttu-id="a750e-127">Pour plus de détails sur l’API, voir [API JavaScript pour PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="a750e-127">For details about the API, see [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="a750e-128">Vérification de la prise en charge d’une exigence d'exécution</span><span class="sxs-lookup"><span data-stu-id="a750e-128">Runtime requirement support check</span></span>

<span data-ttu-id="a750e-129">Lors de l’exécution, les compléments peuvent vérifier si un hôte particulier prend en charge une série de conditions requises d’API en procédant comme suit.</span><span class="sxs-lookup"><span data-stu-id="a750e-129">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="a750e-130">Vérification de la prise en charge des conditions requises basées sur le manifeste</span><span class="sxs-lookup"><span data-stu-id="a750e-130">Manifest-based requirement support check</span></span>

<span data-ttu-id="a750e-131">Utilisez l’élément `Requirements` dans le manifeste du complément pour spécifier des ensembles de conditions requises essentiels ou des membres d’API que votre complément doit utiliser.</span><span class="sxs-lookup"><span data-stu-id="a750e-131">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="a750e-132">Si la plateforme ou l’hôte Office ne prend pas en charge les ensembles de conditions requises ou les membres d’API spécifiés dans l’élément `Requirements`, le complément ne s’exécute pas dans cet hôte ou cette plateforme et ne s’affiche pas dans Mes compléments.</span><span class="sxs-lookup"><span data-stu-id="a750e-132">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="a750e-133">Cet exemple de code illustre un complément qui se charge dans toutes les applications hôtes Office qui prennent en charge l’ensemble de conditions requises OneNoteApi, version 1.1.</span><span class="sxs-lookup"><span data-stu-id="a750e-133">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="a750e-134">Séries de conditions requises des API communes pour Office</span><span class="sxs-lookup"><span data-stu-id="a750e-134">Office Common API requirement sets</span></span>

<span data-ttu-id="a750e-135">La plupart des fonctionnalités du complément PowerPoint proviennent de la série courante d’API.</span><span class="sxs-lookup"><span data-stu-id="a750e-135">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="a750e-136">Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="a750e-136">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="a750e-137">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a750e-137">See also</span></span>

- [<span data-ttu-id="a750e-138">Documentation référence de l’API JavaScript pour PowerPoint</span><span class="sxs-lookup"><span data-stu-id="a750e-138">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="a750e-139">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="a750e-139">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="a750e-140">Spécification des exigences en matière d’hôtes Office et d’API</span><span class="sxs-lookup"><span data-stu-id="a750e-140">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="a750e-141">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="a750e-141">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
