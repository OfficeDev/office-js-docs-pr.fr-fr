---
title: Ensembles de conditions requises de l’API d’identité
description: Informations sur les conditions requises de l’API Identity pour les compléments Office.
ms.date: 07/30/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 05805451f17cc70597a61e55d1ecacbb81c383c5
ms.sourcegitcommit: 8fdd7369bfd97a273e222a0404e337ba2b8807b0
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/05/2020
ms.locfileid: "46573216"
---
# <a name="identity-api-requirement-sets"></a><span data-ttu-id="e2028-103">Ensembles de conditions requises de l’API d’identité</span><span class="sxs-lookup"><span data-stu-id="e2028-103">Identity API requirement sets</span></span>

<span data-ttu-id="e2028-p101">Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="e2028-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="e2028-107">Les compléments Office s’exécutent sur plusieurs versions d’Office.</span><span class="sxs-lookup"><span data-stu-id="e2028-107">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="e2028-108">Le tableau suivant répertorie les ensembles de conditions requises de l’API de boîte de dialogue, les applications Office hôtes qui prennent en charge ces conditions et les numéros de build ou de version de l’application Office.</span><span class="sxs-lookup"><span data-stu-id="e2028-108">The following table lists the Identity API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="e2028-109">Ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="e2028-109">Requirement set</span></span>  | <span data-ttu-id="e2028-110">Office 2013 ou version ultérieure sous Windows</span><span class="sxs-lookup"><span data-stu-id="e2028-110">Office 2013 or later on Windows</span></span><br><span data-ttu-id="e2028-111">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="e2028-111">(one-time purchase)</span></span> | <span data-ttu-id="e2028-112">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="e2028-112">Office on Windows</span></span><br><span data-ttu-id="e2028-113">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="e2028-113">(connected to a Microsoft 365 subscription)</span></span> |  <span data-ttu-id="e2028-114">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="e2028-114">Office on iPad</span></span><br><span data-ttu-id="e2028-115">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="e2028-115">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="e2028-116">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="e2028-116">Office on Mac</span></span><br><span data-ttu-id="e2028-117">(connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="e2028-117">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="e2028-118">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="e2028-118">Office on the web</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="e2028-119">Ensembles 1,3</span><span class="sxs-lookup"><span data-stu-id="e2028-119">IdentityAPI 1.3</span></span>  | <span data-ttu-id="e2028-120">S/O</span><span class="sxs-lookup"><span data-stu-id="e2028-120">N/A</span></span> | <span data-ttu-id="e2028-121">2008 (Build 13127,20000) ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="e2028-121">2008 (build 13127.20000) or later</span></span> | <span data-ttu-id="e2028-122">Bientôt disponible</span><span class="sxs-lookup"><span data-stu-id="e2028-122">Coming soon</span></span> | <span data-ttu-id="e2028-123">16,40 ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="e2028-123">16.40 or later</span></span> | <span data-ttu-id="e2028-124">Août, 2020 \*</span><span class="sxs-lookup"><span data-stu-id="e2028-124">August, 2020\*</span></span> |

> <span data-ttu-id="e2028-125">\*Initialement, l’ensemble de conditions requises est pris en charge dans Office sur le Web uniquement pour les documents ouverts à partir de SharePoint Online et OneDrive.com.</span><span class="sxs-lookup"><span data-stu-id="e2028-125">\* Initially, the requirement set is supported in Office on the web only for documents that are opened from SharePoint Online and OneDrive.com.</span></span> <span data-ttu-id="e2028-126">La prise en charge d’autres documents arrivera sur Office sur le Web plus tard dans 2020.</span><span class="sxs-lookup"><span data-stu-id="e2028-126">Support for other documents will come to Office on the web later in 2020.</span></span>

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="e2028-127">Numéros de version et de build d’Office</span><span class="sxs-lookup"><span data-stu-id="e2028-127">Office versions and build numbers</span></span>

<span data-ttu-id="e2028-128">Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :</span><span class="sxs-lookup"><span data-stu-id="e2028-128">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [<span data-ttu-id="e2028-129">Présentation d’Office Online Server</span><span class="sxs-lookup"><span data-stu-id="e2028-129">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="e2028-130">Ensembles de conditions requises des API communes pour Office</span><span class="sxs-lookup"><span data-stu-id="e2028-130">Office Common API requirement sets</span></span>

<span data-ttu-id="e2028-131">Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="e2028-131">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="identityapi-preview"></a><span data-ttu-id="e2028-132">Préversion ensembles</span><span class="sxs-lookup"><span data-stu-id="e2028-132">IdentityAPI Preview</span></span>

<span data-ttu-id="e2028-133">Pour plus d’informations sur cette API, consultez la version qui utilise les promesses sur [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) ou la version qui utilise les rappels sur [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-).</span><span class="sxs-lookup"><span data-stu-id="e2028-133">For details about this API, see either the version that uses Promises at [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) or the version that uses callbacks at [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-).</span></span>

## <a name="see-also"></a><span data-ttu-id="e2028-134">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e2028-134">See also</span></span>

- [<span data-ttu-id="e2028-135">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="e2028-135">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="e2028-136">Spécification des exigences en matière d’hôtes Office et d’API</span><span class="sxs-lookup"><span data-stu-id="e2028-136">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="e2028-137">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="e2028-137">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
