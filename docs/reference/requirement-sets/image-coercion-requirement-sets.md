---
title: Ensembles de conditions requises de coercition d’image
description: Prise en charge des ensembles de conditions requises pour le foragage d’image avec Office pour les Excel, PowerPoint et Word.
ms.date: 02/19/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 29614718378fd51013360a2a922e11f89bca14b8
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350217"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="37201-103">Ensembles de conditions requises de coercition d’image</span><span class="sxs-lookup"><span data-stu-id="37201-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="37201-p101">Les ensembles de conditions requises sont des groupes nommés des membres de l’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si une application Office prend en charge les API requises par un complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="37201-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="37201-107">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="37201-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="37201-108">ImageCoercion 1.1 permet la conversion en image ( ) lors de l’écriture de données `Office.CoercionType.Image` à l’aide de la [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) méthode.</span><span class="sxs-lookup"><span data-stu-id="37201-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="37201-109">Les applications suivantes sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="37201-109">The following applications are supported.</span></span>

- <span data-ttu-id="37201-110">Excel 2013 et les ultérieures Windows</span><span class="sxs-lookup"><span data-stu-id="37201-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="37201-111">Excel 2016 et ultérieures sur Mac</span><span class="sxs-lookup"><span data-stu-id="37201-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="37201-112">Excel sur iPad</span><span class="sxs-lookup"><span data-stu-id="37201-112">Excel on iPad</span></span>
- <span data-ttu-id="37201-113">OneNote sur le web</span><span class="sxs-lookup"><span data-stu-id="37201-113">OneNote on the web</span></span>
- <span data-ttu-id="37201-114">PowerPoint 2013 et les Windows</span><span class="sxs-lookup"><span data-stu-id="37201-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="37201-115">PowerPoint 2016 et ultérieures sur Mac</span><span class="sxs-lookup"><span data-stu-id="37201-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="37201-116">PowerPoint sur le web</span><span class="sxs-lookup"><span data-stu-id="37201-116">PowerPoint on the web</span></span>
- <span data-ttu-id="37201-117">PowerPoint sur iPad</span><span class="sxs-lookup"><span data-stu-id="37201-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="37201-118">Word 2013 ou version ultérieure sur Windows</span><span class="sxs-lookup"><span data-stu-id="37201-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="37201-119">Word 2016 ou version ultérieure sur Mac</span><span class="sxs-lookup"><span data-stu-id="37201-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="37201-120">Word sur le web</span><span class="sxs-lookup"><span data-stu-id="37201-120">Word on the web</span></span>
- <span data-ttu-id="37201-121">Word sur iPad</span><span class="sxs-lookup"><span data-stu-id="37201-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="37201-122">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="37201-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="37201-123">ImageCoercion 1.2 permet la conversion au format SVG () lors de l’écriture de données `Office.CoercionType.XmlSvg` à l’aide de la [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) méthode.</span><span class="sxs-lookup"><span data-stu-id="37201-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="37201-124">Les applications suivantes sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="37201-124">The following applications are supported.</span></span>

- <span data-ttu-id="37201-125">Excel sur Windows (connecté à un abonnement Microsoft 365 abonnement)</span><span class="sxs-lookup"><span data-stu-id="37201-125">Excel on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="37201-126">Excel mac (connecté à un abonnement Microsoft 365))</span><span class="sxs-lookup"><span data-stu-id="37201-126">Excel on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="37201-127">PowerPoint sur Windows (connecté à un abonnement Microsoft 365 abonnement)</span><span class="sxs-lookup"><span data-stu-id="37201-127">PowerPoint on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="37201-128">PowerPoint mac (connecté à un abonnement Microsoft 365))</span><span class="sxs-lookup"><span data-stu-id="37201-128">PowerPoint on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="37201-129">PowerPoint sur le web</span><span class="sxs-lookup"><span data-stu-id="37201-129">PowerPoint on the web</span></span>
- <span data-ttu-id="37201-130">Word on Windows (connecté à un abonnement Microsoft 365))</span><span class="sxs-lookup"><span data-stu-id="37201-130">Word on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="37201-131">Word sur Mac (connecté à un abonnement Microsoft 365))</span><span class="sxs-lookup"><span data-stu-id="37201-131">Word on Mac (connected to a Microsoft 365 subscription)</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="37201-132">Séries de conditions requises des API communes pour Office</span><span class="sxs-lookup"><span data-stu-id="37201-132">Office Common API requirement sets</span></span>

<span data-ttu-id="37201-133">Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="37201-133">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="37201-134">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="37201-134">See also</span></span>

- [<span data-ttu-id="37201-135">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="37201-135">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="37201-136">Spécifier les exigences en matière d’applications Office et d’API</span><span class="sxs-lookup"><span data-stu-id="37201-136">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="37201-137">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="37201-137">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
