---
title: Ensembles de conditions requises de coercition d’image
description: Prise en charge des ensembles de conditions requises de forçage d’image avec des compléments Office dans Excel, PowerPoint et Word.
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 7140099757c6e4b5ad405723d5fed95fded6d919
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293547"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="4a6e6-103">Ensembles de conditions requises de coercition d’image</span><span class="sxs-lookup"><span data-stu-id="4a6e6-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="4a6e6-104">Les ensembles de conditions requises sont des groupes nommés de membres d’API.</span><span class="sxs-lookup"><span data-stu-id="4a6e6-104">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="4a6e6-105">Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification à l’exécution pour déterminer si une application Office prend en charge les API dont un complément a besoin.</span><span class="sxs-lookup"><span data-stu-id="4a6e6-105">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs.</span></span> <span data-ttu-id="4a6e6-106">Pour plus d’informations, consultez la rubrique [versions d’Office et ensembles de conditions requises](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="4a6e6-106">For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="4a6e6-107">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="4a6e6-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="4a6e6-108">ImageCoercion 1,1 permet la conversion en image ( `Office.CoercionType.Image` ) lors de l’écriture de données à l’aide de la [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) méthode.</span><span class="sxs-lookup"><span data-stu-id="4a6e6-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="4a6e6-109">Les applications suivantes sont prises en charge :</span><span class="sxs-lookup"><span data-stu-id="4a6e6-109">The following applications are supported:</span></span>

- <span data-ttu-id="4a6e6-110">Excel 2013 et versions ultérieures sur Windows</span><span class="sxs-lookup"><span data-stu-id="4a6e6-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="4a6e6-111">Excel 2016 et versions ultérieures sur Mac</span><span class="sxs-lookup"><span data-stu-id="4a6e6-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="4a6e6-112">Excel sur iPad</span><span class="sxs-lookup"><span data-stu-id="4a6e6-112">Excel on iPad</span></span>
- <span data-ttu-id="4a6e6-113">OneNote sur le web</span><span class="sxs-lookup"><span data-stu-id="4a6e6-113">OneNote on the web</span></span>
- <span data-ttu-id="4a6e6-114">PowerPoint 2013 et versions ultérieures sur Windows</span><span class="sxs-lookup"><span data-stu-id="4a6e6-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="4a6e6-115">PowerPoint 2016 et versions ultérieures sur Mac</span><span class="sxs-lookup"><span data-stu-id="4a6e6-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="4a6e6-116">PowerPoint sur le web</span><span class="sxs-lookup"><span data-stu-id="4a6e6-116">PowerPoint on the web</span></span>
- <span data-ttu-id="4a6e6-117">PowerPoint sur iPad</span><span class="sxs-lookup"><span data-stu-id="4a6e6-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="4a6e6-118">Word 2013 ou version ultérieure sur Windows</span><span class="sxs-lookup"><span data-stu-id="4a6e6-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="4a6e6-119">Word 2016 ou version ultérieure sur Mac</span><span class="sxs-lookup"><span data-stu-id="4a6e6-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="4a6e6-120">Word sur le web</span><span class="sxs-lookup"><span data-stu-id="4a6e6-120">Word on the web</span></span>
- <span data-ttu-id="4a6e6-121">Word sur iPad</span><span class="sxs-lookup"><span data-stu-id="4a6e6-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="4a6e6-122">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="4a6e6-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="4a6e6-123">ImageCoercion 1,2 permet d’effectuer une conversion au format SVG ( `Office.CoercionType.XmlSvg` ) lors de l’écriture de données à l’aide de la [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) méthode.</span><span class="sxs-lookup"><span data-stu-id="4a6e6-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="4a6e6-124">Les applications suivantes sont prises en charge :</span><span class="sxs-lookup"><span data-stu-id="4a6e6-124">The following applications are supported:</span></span>

- <span data-ttu-id="4a6e6-125">Excel sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="4a6e6-125">Excel on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="4a6e6-126">Excel sur Mac (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="4a6e6-126">Excel on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="4a6e6-127">PowerPoint sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="4a6e6-127">PowerPoint on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="4a6e6-128">PowerPoint sur Mac (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="4a6e6-128">PowerPoint on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="4a6e6-129">PowerPoint sur le web</span><span class="sxs-lookup"><span data-stu-id="4a6e6-129">PowerPoint on the web</span></span>
- <span data-ttu-id="4a6e6-130">Word sur Windows (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="4a6e6-130">Word on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="4a6e6-131">Word sur Mac (connecté à un abonnement Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="4a6e6-131">Word on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="4a6e6-132">Word sur le web</span><span class="sxs-lookup"><span data-stu-id="4a6e6-132">Word on the web</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="4a6e6-133">Ensembles de conditions requises des API communes pour Office</span><span class="sxs-lookup"><span data-stu-id="4a6e6-133">Office Common API requirement sets</span></span>

<span data-ttu-id="4a6e6-134">Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="4a6e6-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="4a6e6-135">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="4a6e6-135">See also</span></span>

- [<span data-ttu-id="4a6e6-136">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="4a6e6-136">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="4a6e6-137">Spécification des exigences en matière d’applications et d’API Office</span><span class="sxs-lookup"><span data-stu-id="4a6e6-137">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="4a6e6-138">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="4a6e6-138">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
