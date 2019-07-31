---
title: Ensembles de conditions requises de forçage d’image
description: Prise en charge des ensembles de conditions requises de forçage d’image avec des compléments Office dans Excel, PowerPoint et Word.
ms.date: 07/11/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: bffe6c074d9e0734299d0087f2488524875931ed
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940844"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="f17ce-103">Ensembles de conditions requises de forçage d’image</span><span class="sxs-lookup"><span data-stu-id="f17ce-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="f17ce-p101">Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="f17ce-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="f17ce-107">Les compléments Office s’exécutent sur plusieurs versions d’Office.</span><span class="sxs-lookup"><span data-stu-id="f17ce-107">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="f17ce-108">Le tableau suivant répertorie les ensembles de conditions requises de forçage d’image, les applications hôtes Office qui prennent en charge l’ensemble de conditions requises, ainsi que les numéros de build ou de version de l’application Office.</span><span class="sxs-lookup"><span data-stu-id="f17ce-108">The following table lists the Image Coercion requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="f17ce-109">ImageCoercion 1,1</span><span class="sxs-lookup"><span data-stu-id="f17ce-109">ImageCoercion 1.1</span></span>

<span data-ttu-id="f17ce-110">ImageCoercion 1,1 permet la conversion en image (`Office.CoercionType.Image`) lors de l’écriture de [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) données à l’aide de la méthode.</span><span class="sxs-lookup"><span data-stu-id="f17ce-110">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="f17ce-111">Les hôtes suivants sont pris en charge:</span><span class="sxs-lookup"><span data-stu-id="f17ce-111">The following hosts are supported:</span></span>

- <span data-ttu-id="f17ce-112">Excel 2013 et versions ultérieures sur Windows</span><span class="sxs-lookup"><span data-stu-id="f17ce-112">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="f17ce-113">Excel 2016 et versions ultérieures sur Mac</span><span class="sxs-lookup"><span data-stu-id="f17ce-113">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="f17ce-114">Excel sur le Web</span><span class="sxs-lookup"><span data-stu-id="f17ce-114">Excel on the web</span></span>
- <span data-ttu-id="f17ce-115">Excel sur iPad</span><span class="sxs-lookup"><span data-stu-id="f17ce-115">Excel on iPad</span></span>
- <span data-ttu-id="f17ce-116">OneNote sur le Web</span><span class="sxs-lookup"><span data-stu-id="f17ce-116">OneNote on the web</span></span>
- <span data-ttu-id="f17ce-117">PowerPoint 2013 et versions ultérieures sur Windows</span><span class="sxs-lookup"><span data-stu-id="f17ce-117">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="f17ce-118">PowerPoint 2016 et versions ultérieures sur Mac</span><span class="sxs-lookup"><span data-stu-id="f17ce-118">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="f17ce-119">PowerPoint sur le Web</span><span class="sxs-lookup"><span data-stu-id="f17ce-119">PowerPoint on the web</span></span>
- <span data-ttu-id="f17ce-120">PowerPoint sur iPad</span><span class="sxs-lookup"><span data-stu-id="f17ce-120">PowerPoint on iPad</span></span>
- <span data-ttu-id="f17ce-121">Word 2013 et versions ultérieures sur Windows</span><span class="sxs-lookup"><span data-stu-id="f17ce-121">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="f17ce-122">Word 2016 et versions ultérieures sur Mac</span><span class="sxs-lookup"><span data-stu-id="f17ce-122">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="f17ce-123">Word sur le Web</span><span class="sxs-lookup"><span data-stu-id="f17ce-123">Word on the web</span></span>
- <span data-ttu-id="f17ce-124">Word pour iPad</span><span class="sxs-lookup"><span data-stu-id="f17ce-124">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="f17ce-125">ImageCoercion 1,2</span><span class="sxs-lookup"><span data-stu-id="f17ce-125">ImageCoercion 1.2</span></span>

<span data-ttu-id="f17ce-126">ImageCoercion 1,2 permet d’effectuer une conversion au`Office.CoercionType.XmlSvg`format SVG () lors de [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) l’écriture de données à l’aide de la méthode.</span><span class="sxs-lookup"><span data-stu-id="f17ce-126">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="f17ce-127">Les hôtes suivants sont pris en charge:</span><span class="sxs-lookup"><span data-stu-id="f17ce-127">The following hosts are supported:</span></span>

- <span data-ttu-id="f17ce-128">Excel sur Windows (connecté à un abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f17ce-128">Excel on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="f17ce-129">Excel sur Mac (connecté à un abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f17ce-129">Excel on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="f17ce-130">Excel sur le Web</span><span class="sxs-lookup"><span data-stu-id="f17ce-130">Excel on the web</span></span>
- <span data-ttu-id="f17ce-131">PowerPoint sur Windows (connecté à un abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f17ce-131">PowerPoint on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="f17ce-132">PowerPoint sur Mac (connecté à un abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f17ce-132">PowerPoint on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="f17ce-133">PowerPoint sur le Web</span><span class="sxs-lookup"><span data-stu-id="f17ce-133">PowerPoint on the web</span></span>
- <span data-ttu-id="f17ce-134">Word sur Windows (connecté à un abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f17ce-134">Word on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="f17ce-135">Word sur Mac (connecté à un abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="f17ce-135">Word on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="f17ce-136">Word sur le Web</span><span class="sxs-lookup"><span data-stu-id="f17ce-136">Word on the web</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="f17ce-137">Ensembles de conditions requises des API communes pour Office</span><span class="sxs-lookup"><span data-stu-id="f17ce-137">Office Common API requirement sets</span></span>

<span data-ttu-id="f17ce-138">Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="f17ce-138">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="f17ce-139">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f17ce-139">See also</span></span>

- [<span data-ttu-id="f17ce-140">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="f17ce-140">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="f17ce-141">Spécification des exigences en matière d’hôtes Office et d’API</span><span class="sxs-lookup"><span data-stu-id="f17ce-141">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="f17ce-142">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="f17ce-142">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
