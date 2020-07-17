---
title: Ensembles de conditions requises de coercition d’image
description: Prise en charge des ensembles de conditions requises de forçage d’image avec des compléments Office dans Excel, PowerPoint et Word.
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 59f6891182f47bed1b7e3b6aa69a30e941bce7cb
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094352"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="e3f14-103">Ensembles de conditions requises de coercition d’image</span><span class="sxs-lookup"><span data-stu-id="e3f14-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="e3f14-p101">Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="e3f14-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="e3f14-107">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="e3f14-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="e3f14-108">ImageCoercion 1,1 permet la conversion en image ( `Office.CoercionType.Image` ) lors de l’écriture de données à l’aide de la [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) méthode.</span><span class="sxs-lookup"><span data-stu-id="e3f14-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="e3f14-109">Les hôtes suivants sont pris en charge :</span><span class="sxs-lookup"><span data-stu-id="e3f14-109">The following hosts are supported:</span></span>

- <span data-ttu-id="e3f14-110">Excel 2013 et versions ultérieures sur Windows</span><span class="sxs-lookup"><span data-stu-id="e3f14-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="e3f14-111">Excel 2016 et versions ultérieures sur Mac</span><span class="sxs-lookup"><span data-stu-id="e3f14-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="e3f14-112">Excel sur iPad</span><span class="sxs-lookup"><span data-stu-id="e3f14-112">Excel on iPad</span></span>
- <span data-ttu-id="e3f14-113">OneNote sur le web</span><span class="sxs-lookup"><span data-stu-id="e3f14-113">OneNote on the web</span></span>
- <span data-ttu-id="e3f14-114">PowerPoint 2013 et versions ultérieures sur Windows</span><span class="sxs-lookup"><span data-stu-id="e3f14-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="e3f14-115">PowerPoint 2016 et versions ultérieures sur Mac</span><span class="sxs-lookup"><span data-stu-id="e3f14-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="e3f14-116">PowerPoint sur le web</span><span class="sxs-lookup"><span data-stu-id="e3f14-116">PowerPoint on the web</span></span>
- <span data-ttu-id="e3f14-117">PowerPoint sur iPad</span><span class="sxs-lookup"><span data-stu-id="e3f14-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="e3f14-118">Word 2013 ou version ultérieure sur Windows</span><span class="sxs-lookup"><span data-stu-id="e3f14-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="e3f14-119">Word 2016 ou version ultérieure sur Mac</span><span class="sxs-lookup"><span data-stu-id="e3f14-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="e3f14-120">Word sur le web</span><span class="sxs-lookup"><span data-stu-id="e3f14-120">Word on the web</span></span>
- <span data-ttu-id="e3f14-121">Word sur iPad</span><span class="sxs-lookup"><span data-stu-id="e3f14-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="e3f14-122">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="e3f14-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="e3f14-123">ImageCoercion 1,2 permet d’effectuer une conversion au format SVG ( `Office.CoercionType.XmlSvg` ) lors de l’écriture de données à l’aide de la [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) méthode.</span><span class="sxs-lookup"><span data-stu-id="e3f14-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="e3f14-124">Les hôtes suivants sont pris en charge :</span><span class="sxs-lookup"><span data-stu-id="e3f14-124">The following hosts are supported:</span></span>

- <span data-ttu-id="e3f14-125">Excel sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="e3f14-125">Excel on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="e3f14-126">Excel sur Mac (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="e3f14-126">Excel on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="e3f14-127">PowerPoint sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="e3f14-127">PowerPoint on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="e3f14-128">PowerPoint sur Mac (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="e3f14-128">PowerPoint on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="e3f14-129">PowerPoint sur le web</span><span class="sxs-lookup"><span data-stu-id="e3f14-129">PowerPoint on the web</span></span>
- <span data-ttu-id="e3f14-130">Word sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="e3f14-130">Word on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="e3f14-131">Word sur Mac (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="e3f14-131">Word on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="e3f14-132">Word sur le web</span><span class="sxs-lookup"><span data-stu-id="e3f14-132">Word on the web</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="e3f14-133">Ensembles de conditions requises des API communes pour Office</span><span class="sxs-lookup"><span data-stu-id="e3f14-133">Office Common API requirement sets</span></span>

<span data-ttu-id="e3f14-134">Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="e3f14-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="e3f14-135">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e3f14-135">See also</span></span>

- [<span data-ttu-id="e3f14-136">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3f14-136">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="e3f14-137">Spécification des exigences en matière d’hôtes Office et d’API</span><span class="sxs-lookup"><span data-stu-id="e3f14-137">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="e3f14-138">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="e3f14-138">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
