---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, OneNote, Outlook, PowerPoint, Project et Word.
ms.date: 07/11/2019
localization_priority: Priority
ms.openlocfilehash: 2bfeb7cc5c6e8846f1d882abf3a0149302e53914
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771834"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="73e8e-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="73e8e-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="73e8e-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="73e8e-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="73e8e-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="73e8e-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="73e8e-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="73e8e-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="73e8e-108">Excel</span><span class="sxs-lookup"><span data-stu-id="73e8e-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="73e8e-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="73e8e-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="73e8e-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="73e8e-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="73e8e-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="73e8e-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="73e8e-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="73e8e-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-113">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="73e8e-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="73e8e-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-114">- TaskPane</span></span><br><span data-ttu-id="73e8e-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="73e8e-115">
        - Content</span></span><br><span data-ttu-id="73e8e-116">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="73e8e-116">
        - Custom Functions</span></span><br><span data-ttu-id="73e8e-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="73e8e-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="73e8e-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="73e8e-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="73e8e-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="73e8e-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="73e8e-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="73e8e-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="73e8e-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="73e8e-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="73e8e-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="73e8e-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="73e8e-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="73e8e-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="73e8e-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-130">
        - BindingEvents</span></span><br><span data-ttu-id="73e8e-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-131">
        - CompressedFile</span></span><br><span data-ttu-id="73e8e-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-132">
        - DocumentEvents</span></span><br><span data-ttu-id="73e8e-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-133">
        - File</span></span><br><span data-ttu-id="73e8e-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-134">
        - MatrixBindings</span></span><br><span data-ttu-id="73e8e-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="73e8e-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-136">
        - Selection</span></span><br><span data-ttu-id="73e8e-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-137">
        - Settings</span></span><br><span data-ttu-id="73e8e-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-138">
        - TableBindings</span></span><br><span data-ttu-id="73e8e-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-139">
        - TableCoercion</span></span><br><span data-ttu-id="73e8e-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-140">
        - TextBindings</span></span><br><span data-ttu-id="73e8e-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-142">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-142">Office on Windows</span></span><br><span data-ttu-id="73e8e-143">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="73e8e-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="73e8e-144">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-144">- TaskPane</span></span><br><span data-ttu-id="73e8e-145">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="73e8e-145">
        - Content</span></span><br><span data-ttu-id="73e8e-146">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="73e8e-146">
        - Custom Functions</span></span><br><span data-ttu-id="73e8e-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="73e8e-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="73e8e-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="73e8e-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="73e8e-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="73e8e-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="73e8e-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="73e8e-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="73e8e-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="73e8e-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="73e8e-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="73e8e-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="73e8e-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="73e8e-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="73e8e-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-160">
        - BindingEvents</span></span><br><span data-ttu-id="73e8e-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-161">
        - CompressedFile</span></span><br><span data-ttu-id="73e8e-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-162">
        - DocumentEvents</span></span><br><span data-ttu-id="73e8e-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-163">
        - File</span></span><br><span data-ttu-id="73e8e-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-164">
        - MatrixBindings</span></span><br><span data-ttu-id="73e8e-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="73e8e-166">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-166">
        - Selection</span></span><br><span data-ttu-id="73e8e-167">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-167">
        - Settings</span></span><br><span data-ttu-id="73e8e-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-168">
        - TableBindings</span></span><br><span data-ttu-id="73e8e-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-169">
        - TableCoercion</span></span><br><span data-ttu-id="73e8e-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-170">
        - TextBindings</span></span><br><span data-ttu-id="73e8e-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-172">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-172">Office 2019 on Windows</span></span><br><span data-ttu-id="73e8e-173">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="73e8e-174">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-174">- TaskPane</span></span><br><span data-ttu-id="73e8e-175">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="73e8e-175">
        - Content</span></span><br><span data-ttu-id="73e8e-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="73e8e-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="73e8e-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="73e8e-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="73e8e-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="73e8e-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="73e8e-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="73e8e-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="73e8e-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="73e8e-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="73e8e-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="73e8e-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-187">- BindingEvents</span></span><br><span data-ttu-id="73e8e-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-188">
        - CompressedFile</span></span><br><span data-ttu-id="73e8e-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-189">
        - DocumentEvents</span></span><br><span data-ttu-id="73e8e-190">
        - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-190">
        - File</span></span><br><span data-ttu-id="73e8e-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-191">
        - MatrixBindings</span></span><br><span data-ttu-id="73e8e-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="73e8e-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-193">
        - Selection</span></span><br><span data-ttu-id="73e8e-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-194">
        - Settings</span></span><br><span data-ttu-id="73e8e-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-195">
        - TableBindings</span></span><br><span data-ttu-id="73e8e-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-196">
        - TableCoercion</span></span><br><span data-ttu-id="73e8e-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-197">
        - TextBindings</span></span><br><span data-ttu-id="73e8e-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-199">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-199">Office 2016 on Windows</span></span><br><span data-ttu-id="73e8e-200">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="73e8e-201">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-201">- TaskPane</span></span><br><span data-ttu-id="73e8e-202">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="73e8e-202">
        - Content</span></span></td>
    <td><span data-ttu-id="73e8e-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="73e8e-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="73e8e-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="73e8e-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="73e8e-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-206">- BindingEvents</span></span><br><span data-ttu-id="73e8e-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-207">
        - CompressedFile</span></span><br><span data-ttu-id="73e8e-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-208">
        - DocumentEvents</span></span><br><span data-ttu-id="73e8e-209">
        - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-209">
        - File</span></span><br><span data-ttu-id="73e8e-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-210">
        - MatrixBindings</span></span><br><span data-ttu-id="73e8e-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="73e8e-212">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-212">
        - Selection</span></span><br><span data-ttu-id="73e8e-213">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-213">
        - Settings</span></span><br><span data-ttu-id="73e8e-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-214">
        - TableBindings</span></span><br><span data-ttu-id="73e8e-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-215">
        - TableCoercion</span></span><br><span data-ttu-id="73e8e-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-216">
        - TextBindings</span></span><br><span data-ttu-id="73e8e-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-218">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-218">Office 2013 on Windows</span></span><br><span data-ttu-id="73e8e-219">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="73e8e-220">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-220">
        - TaskPane</span></span><br><span data-ttu-id="73e8e-221">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="73e8e-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="73e8e-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="73e8e-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="73e8e-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="73e8e-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-224">
        - BindingEvents</span></span><br><span data-ttu-id="73e8e-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-225">
        - CompressedFile</span></span><br><span data-ttu-id="73e8e-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-226">
        - DocumentEvents</span></span><br><span data-ttu-id="73e8e-227">
        - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-227">
        - File</span></span><br><span data-ttu-id="73e8e-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-228">
        - MatrixBindings</span></span><br><span data-ttu-id="73e8e-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="73e8e-230">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-230">
        - Selection</span></span><br><span data-ttu-id="73e8e-231">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-231">
        - Settings</span></span><br><span data-ttu-id="73e8e-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-232">
        - TableBindings</span></span><br><span data-ttu-id="73e8e-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-233">
        - TableCoercion</span></span><br><span data-ttu-id="73e8e-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-234">
        - TextBindings</span></span><br><span data-ttu-id="73e8e-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-236">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="73e8e-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="73e8e-237">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="73e8e-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="73e8e-238">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-238">- TaskPane</span></span><br><span data-ttu-id="73e8e-239">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="73e8e-239">
        - Content</span></span><br><span data-ttu-id="73e8e-240">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="73e8e-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="73e8e-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="73e8e-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="73e8e-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="73e8e-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="73e8e-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="73e8e-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="73e8e-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="73e8e-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="73e8e-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="73e8e-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="73e8e-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="73e8e-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-252">- BindingEvents</span></span><br><span data-ttu-id="73e8e-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-253">
        - DocumentEvents</span></span><br><span data-ttu-id="73e8e-254">
        - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-254">
        - File</span></span><br><span data-ttu-id="73e8e-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-255">
        - MatrixBindings</span></span><br><span data-ttu-id="73e8e-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="73e8e-257">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-257">
        - Selection</span></span><br><span data-ttu-id="73e8e-258">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-258">
        - Settings</span></span><br><span data-ttu-id="73e8e-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-259">
        - TableBindings</span></span><br><span data-ttu-id="73e8e-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-260">
        - TableCoercion</span></span><br><span data-ttu-id="73e8e-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-261">
        - TextBindings</span></span><br><span data-ttu-id="73e8e-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-263">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="73e8e-263">Office apps on Mac</span></span><br><span data-ttu-id="73e8e-264">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="73e8e-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="73e8e-265">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-265">- TaskPane</span></span><br><span data-ttu-id="73e8e-266">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="73e8e-266">
        - Content</span></span><br><span data-ttu-id="73e8e-267">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="73e8e-267">
        - Custom Functions</span></span><br><span data-ttu-id="73e8e-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="73e8e-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="73e8e-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="73e8e-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="73e8e-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="73e8e-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="73e8e-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="73e8e-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="73e8e-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="73e8e-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="73e8e-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="73e8e-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="73e8e-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="73e8e-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-281">- BindingEvents</span></span><br><span data-ttu-id="73e8e-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-282">
        - CompressedFile</span></span><br><span data-ttu-id="73e8e-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-283">
        - DocumentEvents</span></span><br><span data-ttu-id="73e8e-284">
        - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-284">
        - File</span></span><br><span data-ttu-id="73e8e-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-285">
        - MatrixBindings</span></span><br><span data-ttu-id="73e8e-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="73e8e-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-287">
        - PdfFile</span></span><br><span data-ttu-id="73e8e-288">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-288">
        - Selection</span></span><br><span data-ttu-id="73e8e-289">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-289">
        - Settings</span></span><br><span data-ttu-id="73e8e-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-290">
        - TableBindings</span></span><br><span data-ttu-id="73e8e-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-291">
        - TableCoercion</span></span><br><span data-ttu-id="73e8e-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-292">
        - TextBindings</span></span><br><span data-ttu-id="73e8e-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-294">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="73e8e-294">Office 2019 for Mac</span></span><br><span data-ttu-id="73e8e-295">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="73e8e-296">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-296">- TaskPane</span></span><br><span data-ttu-id="73e8e-297">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="73e8e-297">
        - Content</span></span><br><span data-ttu-id="73e8e-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="73e8e-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="73e8e-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="73e8e-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="73e8e-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="73e8e-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="73e8e-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="73e8e-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="73e8e-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="73e8e-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="73e8e-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="73e8e-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-309">- BindingEvents</span></span><br><span data-ttu-id="73e8e-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-310">
        - CompressedFile</span></span><br><span data-ttu-id="73e8e-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-311">
        - DocumentEvents</span></span><br><span data-ttu-id="73e8e-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-312">
        - File</span></span><br><span data-ttu-id="73e8e-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-313">
        - MatrixBindings</span></span><br><span data-ttu-id="73e8e-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="73e8e-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-315">
        - PdfFile</span></span><br><span data-ttu-id="73e8e-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-316">
        - Selection</span></span><br><span data-ttu-id="73e8e-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-317">
        - Settings</span></span><br><span data-ttu-id="73e8e-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-318">
        - TableBindings</span></span><br><span data-ttu-id="73e8e-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-319">
        - TableCoercion</span></span><br><span data-ttu-id="73e8e-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-320">
        - TextBindings</span></span><br><span data-ttu-id="73e8e-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-322">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="73e8e-322">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="73e8e-323">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="73e8e-324">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-324">- TaskPane</span></span><br><span data-ttu-id="73e8e-325">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="73e8e-325">
        - Content</span></span></td>
    <td><span data-ttu-id="73e8e-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="73e8e-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="73e8e-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="73e8e-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="73e8e-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-329">- BindingEvents</span></span><br><span data-ttu-id="73e8e-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-330">
        - CompressedFile</span></span><br><span data-ttu-id="73e8e-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-331">
        - DocumentEvents</span></span><br><span data-ttu-id="73e8e-332">
        - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-332">
        - File</span></span><br><span data-ttu-id="73e8e-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-333">
        - MatrixBindings</span></span><br><span data-ttu-id="73e8e-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="73e8e-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-335">
        - PdfFile</span></span><br><span data-ttu-id="73e8e-336">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-336">
        - Selection</span></span><br><span data-ttu-id="73e8e-337">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-337">
        - Settings</span></span><br><span data-ttu-id="73e8e-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-338">
        - TableBindings</span></span><br><span data-ttu-id="73e8e-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-339">
        - TableCoercion</span></span><br><span data-ttu-id="73e8e-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-340">
        - TextBindings</span></span><br><span data-ttu-id="73e8e-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="73e8e-342">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="73e8e-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="73e8e-343">Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="73e8e-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="73e8e-344">Plateforme</span><span class="sxs-lookup"><span data-stu-id="73e8e-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="73e8e-345">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="73e8e-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="73e8e-346">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="73e8e-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="73e8e-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="73e8e-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-348">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="73e8e-348">Office on the web</span></span></td>
    <td><span data-ttu-id="73e8e-349">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="73e8e-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="73e8e-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-351">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-351">Office on Windows</span></span><br><span data-ttu-id="73e8e-352">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="73e8e-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="73e8e-353">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="73e8e-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="73e8e-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-355">Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="73e8e-355">Office for Mac</span></span><br><span data-ttu-id="73e8e-356">(connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="73e8e-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="73e8e-357">
        - Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="73e8e-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="73e8e-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="73e8e-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="73e8e-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="73e8e-360">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="73e8e-360">Platform</span></span></th>
    <th><span data-ttu-id="73e8e-361">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="73e8e-361">Extension points</span></span></th>
    <th><span data-ttu-id="73e8e-362">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="73e8e-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="73e8e-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="73e8e-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-364">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="73e8e-364">Office on the web</span></span><br><span data-ttu-id="73e8e-365">(nouveau)</span><span class="sxs-lookup"><span data-stu-id="73e8e-365">New</span></span></td>
    <td> <span data-ttu-id="73e8e-366">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-366">- Mail Read</span></span><br><span data-ttu-id="73e8e-367">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-367">
      - Mail Compose</span></span><br><span data-ttu-id="73e8e-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="73e8e-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="73e8e-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="73e8e-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="73e8e-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="73e8e-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="73e8e-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="73e8e-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="73e8e-376">Non disponible</span><span class="sxs-lookup"><span data-stu-id="73e8e-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-377">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="73e8e-377">Office on the web</span></span><br><span data-ttu-id="73e8e-378">(classique)</span><span class="sxs-lookup"><span data-stu-id="73e8e-378">Classic.</span></span></td>
    <td> <span data-ttu-id="73e8e-379">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-379">- Mail Read</span></span><br><span data-ttu-id="73e8e-380">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-380">
      - Mail Compose</span></span><br><span data-ttu-id="73e8e-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="73e8e-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="73e8e-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="73e8e-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="73e8e-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="73e8e-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="73e8e-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="73e8e-388">Non disponible</span><span class="sxs-lookup"><span data-stu-id="73e8e-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-389">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-389">Office on Windows</span></span><br><span data-ttu-id="73e8e-390">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="73e8e-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="73e8e-391">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-391">- Mail Read</span></span><br><span data-ttu-id="73e8e-392">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-392">
      - Mail Compose</span></span><br><span data-ttu-id="73e8e-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="73e8e-394">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="73e8e-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="73e8e-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="73e8e-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="73e8e-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="73e8e-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="73e8e-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="73e8e-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="73e8e-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="73e8e-402">Non disponible</span><span class="sxs-lookup"><span data-stu-id="73e8e-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-403">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-403">Office 2019 on Windows</span></span><br><span data-ttu-id="73e8e-404">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="73e8e-405">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-405">- Mail Read</span></span><br><span data-ttu-id="73e8e-406">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-406">
      - Mail Compose</span></span><br><span data-ttu-id="73e8e-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="73e8e-408">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="73e8e-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="73e8e-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="73e8e-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="73e8e-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="73e8e-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="73e8e-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="73e8e-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="73e8e-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="73e8e-416">Non disponible</span><span class="sxs-lookup"><span data-stu-id="73e8e-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-417">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-417">Office 2016 on Windows</span></span><br><span data-ttu-id="73e8e-418">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="73e8e-419">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-419">- Mail Read</span></span><br><span data-ttu-id="73e8e-420">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-420">
      - Mail Compose</span></span><br><span data-ttu-id="73e8e-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="73e8e-422">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="73e8e-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="73e8e-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="73e8e-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="73e8e-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="73e8e-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="73e8e-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="73e8e-427">Non disponible</span><span class="sxs-lookup"><span data-stu-id="73e8e-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-428">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-428">Office 2013 on Windows</span></span><br><span data-ttu-id="73e8e-429">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="73e8e-430">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-430">- Mail Read</span></span><br><span data-ttu-id="73e8e-431">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="73e8e-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="73e8e-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="73e8e-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="73e8e-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="73e8e-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="73e8e-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="73e8e-436">Non disponible</span><span class="sxs-lookup"><span data-stu-id="73e8e-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-437">Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="73e8e-437">Office apps on iOS</span></span><br><span data-ttu-id="73e8e-438">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="73e8e-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="73e8e-439">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-439">- Mail Read</span></span><br><span data-ttu-id="73e8e-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="73e8e-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="73e8e-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="73e8e-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="73e8e-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="73e8e-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="73e8e-446">Non disponible</span><span class="sxs-lookup"><span data-stu-id="73e8e-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-447">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="73e8e-447">Office apps on Mac</span></span><br><span data-ttu-id="73e8e-448">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="73e8e-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="73e8e-449">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-449">- Mail Read</span></span><br><span data-ttu-id="73e8e-450">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-450">
      - Mail Compose</span></span><br><span data-ttu-id="73e8e-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="73e8e-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="73e8e-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="73e8e-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="73e8e-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="73e8e-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="73e8e-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="73e8e-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="73e8e-459">Non disponible</span><span class="sxs-lookup"><span data-stu-id="73e8e-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-460">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="73e8e-460">Office 2019 for Mac</span></span><br><span data-ttu-id="73e8e-461">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="73e8e-462">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-462">- Mail Read</span></span><br><span data-ttu-id="73e8e-463">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-463">
      - Mail Compose</span></span><br><span data-ttu-id="73e8e-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="73e8e-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="73e8e-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="73e8e-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="73e8e-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="73e8e-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="73e8e-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="73e8e-471">Non disponible</span><span class="sxs-lookup"><span data-stu-id="73e8e-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-472">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="73e8e-472">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="73e8e-473">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="73e8e-474">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-474">- Mail Read</span></span><br><span data-ttu-id="73e8e-475">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-475">
      - Mail Compose</span></span><br><span data-ttu-id="73e8e-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="73e8e-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="73e8e-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="73e8e-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="73e8e-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="73e8e-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="73e8e-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="73e8e-483">Non disponible</span><span class="sxs-lookup"><span data-stu-id="73e8e-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-484">Office sur Android</span><span class="sxs-lookup"><span data-stu-id="73e8e-484">Office apps on Android</span></span><br><span data-ttu-id="73e8e-485">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="73e8e-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="73e8e-486">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="73e8e-486">- Mail Read</span></span><br><span data-ttu-id="73e8e-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="73e8e-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="73e8e-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="73e8e-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="73e8e-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="73e8e-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="73e8e-493">Non disponible</span><span class="sxs-lookup"><span data-stu-id="73e8e-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="73e8e-494">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="73e8e-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="73e8e-495">Word</span><span class="sxs-lookup"><span data-stu-id="73e8e-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="73e8e-496">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="73e8e-496">Platform</span></span></th>
    <th><span data-ttu-id="73e8e-497">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="73e8e-497">Extension points</span></span></th>
    <th><span data-ttu-id="73e8e-498">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="73e8e-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="73e8e-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="73e8e-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-500">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="73e8e-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="73e8e-501">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-501">- TaskPane</span></span><br><span data-ttu-id="73e8e-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="73e8e-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="73e8e-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="73e8e-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="73e8e-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="73e8e-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="73e8e-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="73e8e-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-509">- BindingEvents</span></span><br><span data-ttu-id="73e8e-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="73e8e-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="73e8e-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-511">
         - DocumentEvents</span></span><br><span data-ttu-id="73e8e-512">
         - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-512">
         - File</span></span><br><span data-ttu-id="73e8e-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="73e8e-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-514">
         - MatrixBindings</span></span><br><span data-ttu-id="73e8e-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="73e8e-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="73e8e-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-517">
         - PdfFile</span></span><br><span data-ttu-id="73e8e-518">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-518">
         - Selection</span></span><br><span data-ttu-id="73e8e-519">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-519">
         - Settings</span></span><br><span data-ttu-id="73e8e-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-520">
         - TableBindings</span></span><br><span data-ttu-id="73e8e-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-521">
         - TableCoercion</span></span><br><span data-ttu-id="73e8e-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-522">
         - TextBindings</span></span><br><span data-ttu-id="73e8e-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-523">
         - TextCoercion</span></span><br><span data-ttu-id="73e8e-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-525">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-525">Office on Windows</span></span><br><span data-ttu-id="73e8e-526">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="73e8e-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="73e8e-527">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-527">- TaskPane</span></span><br><span data-ttu-id="73e8e-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="73e8e-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="73e8e-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="73e8e-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="73e8e-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="73e8e-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="73e8e-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="73e8e-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-535">- BindingEvents</span></span><br><span data-ttu-id="73e8e-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-536">
         - CompressedFile</span></span><br><span data-ttu-id="73e8e-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="73e8e-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="73e8e-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-538">
         - DocumentEvents</span></span><br><span data-ttu-id="73e8e-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-539">
         - File</span></span><br><span data-ttu-id="73e8e-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="73e8e-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-541">
         - MatrixBindings</span></span><br><span data-ttu-id="73e8e-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="73e8e-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="73e8e-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-544">
         - PdfFile</span></span><br><span data-ttu-id="73e8e-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-545">
         - Selection</span></span><br><span data-ttu-id="73e8e-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-546">
         - Settings</span></span><br><span data-ttu-id="73e8e-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-547">
         - TableBindings</span></span><br><span data-ttu-id="73e8e-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-548">
         - TableCoercion</span></span><br><span data-ttu-id="73e8e-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-549">
         - TextBindings</span></span><br><span data-ttu-id="73e8e-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-550">
         - TextCoercion</span></span><br><span data-ttu-id="73e8e-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-552">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-552">Office 2019 on Windows</span></span><br><span data-ttu-id="73e8e-553">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="73e8e-554">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="73e8e-554">- TaskPane</span></span><br><span data-ttu-id="73e8e-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="73e8e-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="73e8e-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="73e8e-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="73e8e-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="73e8e-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="73e8e-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-561">- BindingEvents</span></span><br><span data-ttu-id="73e8e-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-562">
         - CompressedFile</span></span><br><span data-ttu-id="73e8e-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="73e8e-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="73e8e-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-564">
         - DocumentEvents</span></span><br><span data-ttu-id="73e8e-565">
         - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-565">
         - File</span></span><br><span data-ttu-id="73e8e-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="73e8e-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-567">
         - MatrixBindings</span></span><br><span data-ttu-id="73e8e-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="73e8e-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="73e8e-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-570">
         - PdfFile</span></span><br><span data-ttu-id="73e8e-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-571">
         - Selection</span></span><br><span data-ttu-id="73e8e-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-572">
         - Settings</span></span><br><span data-ttu-id="73e8e-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-573">
         - TableBindings</span></span><br><span data-ttu-id="73e8e-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-574">
         - TableCoercion</span></span><br><span data-ttu-id="73e8e-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-575">
         - TextBindings</span></span><br><span data-ttu-id="73e8e-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-576">
         - TextCoercion</span></span><br><span data-ttu-id="73e8e-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-578">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-578">Office 2016 on Windows</span></span><br><span data-ttu-id="73e8e-579">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="73e8e-580">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="73e8e-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="73e8e-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="73e8e-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="73e8e-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="73e8e-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-584">- BindingEvents</span></span><br><span data-ttu-id="73e8e-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-585">
         - CompressedFile</span></span><br><span data-ttu-id="73e8e-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="73e8e-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="73e8e-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-587">
         - DocumentEvents</span></span><br><span data-ttu-id="73e8e-588">
         - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-588">
         - File</span></span><br><span data-ttu-id="73e8e-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="73e8e-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-590">
         - MatrixBindings</span></span><br><span data-ttu-id="73e8e-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="73e8e-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="73e8e-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-593">
         - PdfFile</span></span><br><span data-ttu-id="73e8e-594">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-594">
         - Selection</span></span><br><span data-ttu-id="73e8e-595">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-595">
         - Settings</span></span><br><span data-ttu-id="73e8e-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-596">
         - TableBindings</span></span><br><span data-ttu-id="73e8e-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-597">
         - TableCoercion</span></span><br><span data-ttu-id="73e8e-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-598">
         - TextBindings</span></span><br><span data-ttu-id="73e8e-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-599">
         - TextCoercion</span></span><br><span data-ttu-id="73e8e-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-601">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-601">Office 2013 on Windows</span></span><br><span data-ttu-id="73e8e-602">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="73e8e-603">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="73e8e-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="73e8e-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="73e8e-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="73e8e-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-606">- BindingEvents</span></span><br><span data-ttu-id="73e8e-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-607">
         - CompressedFile</span></span><br><span data-ttu-id="73e8e-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="73e8e-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="73e8e-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-609">
         - DocumentEvents</span></span><br><span data-ttu-id="73e8e-610">
         - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-610">
         - File</span></span><br><span data-ttu-id="73e8e-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="73e8e-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-612">
         - MatrixBindings</span></span><br><span data-ttu-id="73e8e-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="73e8e-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="73e8e-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-615">
         - PdfFile</span></span><br><span data-ttu-id="73e8e-616">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-616">
         - Selection</span></span><br><span data-ttu-id="73e8e-617">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-617">
         - Settings</span></span><br><span data-ttu-id="73e8e-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-618">
         - TableBindings</span></span><br><span data-ttu-id="73e8e-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-619">
         - TableCoercion</span></span><br><span data-ttu-id="73e8e-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-620">
         - TextBindings</span></span><br><span data-ttu-id="73e8e-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-621">
         - TextCoercion</span></span><br><span data-ttu-id="73e8e-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-623">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="73e8e-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="73e8e-624">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="73e8e-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="73e8e-625">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="73e8e-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="73e8e-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="73e8e-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="73e8e-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="73e8e-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="73e8e-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-631">- BindingEvents</span></span><br><span data-ttu-id="73e8e-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-632">
         - CompressedFile</span></span><br><span data-ttu-id="73e8e-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="73e8e-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="73e8e-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-634">
         - DocumentEvents</span></span><br><span data-ttu-id="73e8e-635">
         - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-635">
         - File</span></span><br><span data-ttu-id="73e8e-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="73e8e-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-637">
         - MatrixBindings</span></span><br><span data-ttu-id="73e8e-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="73e8e-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="73e8e-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-640">
         - PdfFile</span></span><br><span data-ttu-id="73e8e-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-641">
         - Selection</span></span><br><span data-ttu-id="73e8e-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-642">
         - Settings</span></span><br><span data-ttu-id="73e8e-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-643">
         - TableBindings</span></span><br><span data-ttu-id="73e8e-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-644">
         - TableCoercion</span></span><br><span data-ttu-id="73e8e-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-645">
         - TextBindings</span></span><br><span data-ttu-id="73e8e-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-646">
         - TextCoercion</span></span><br><span data-ttu-id="73e8e-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-648">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="73e8e-648">Office apps on Mac</span></span><br><span data-ttu-id="73e8e-649">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="73e8e-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="73e8e-650">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-650">- TaskPane</span></span><br><span data-ttu-id="73e8e-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="73e8e-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="73e8e-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="73e8e-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="73e8e-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="73e8e-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="73e8e-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="73e8e-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-658">- BindingEvents</span></span><br><span data-ttu-id="73e8e-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-659">
         - CompressedFile</span></span><br><span data-ttu-id="73e8e-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="73e8e-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="73e8e-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-661">
         - DocumentEvents</span></span><br><span data-ttu-id="73e8e-662">
         - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-662">
         - File</span></span><br><span data-ttu-id="73e8e-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="73e8e-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-664">
         - MatrixBindings</span></span><br><span data-ttu-id="73e8e-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="73e8e-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="73e8e-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-667">
         - PdfFile</span></span><br><span data-ttu-id="73e8e-668">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-668">
         - Selection</span></span><br><span data-ttu-id="73e8e-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-669">
         - Settings</span></span><br><span data-ttu-id="73e8e-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-670">
         - TableBindings</span></span><br><span data-ttu-id="73e8e-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-671">
         - TableCoercion</span></span><br><span data-ttu-id="73e8e-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-672">
         - TextBindings</span></span><br><span data-ttu-id="73e8e-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-673">
         - TextCoercion</span></span><br><span data-ttu-id="73e8e-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-675">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="73e8e-675">Office 2019 for Mac</span></span><br><span data-ttu-id="73e8e-676">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="73e8e-677">- Volet des tâches</span><span class="sxs-lookup"><span data-stu-id="73e8e-677">- TaskPane</span></span><br><span data-ttu-id="73e8e-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="73e8e-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="73e8e-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="73e8e-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="73e8e-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="73e8e-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="73e8e-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-684">- BindingEvents</span></span><br><span data-ttu-id="73e8e-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-685">
         - CompressedFile</span></span><br><span data-ttu-id="73e8e-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="73e8e-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="73e8e-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-687">
         - DocumentEvents</span></span><br><span data-ttu-id="73e8e-688">
         - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-688">
         - File</span></span><br><span data-ttu-id="73e8e-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="73e8e-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-690">
         - MatrixBindings</span></span><br><span data-ttu-id="73e8e-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="73e8e-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="73e8e-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-693">
         - PdfFile</span></span><br><span data-ttu-id="73e8e-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-694">
         - Selection</span></span><br><span data-ttu-id="73e8e-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-695">
         - Settings</span></span><br><span data-ttu-id="73e8e-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-696">
         - TableBindings</span></span><br><span data-ttu-id="73e8e-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-697">
         - TableCoercion</span></span><br><span data-ttu-id="73e8e-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-698">
         - TextBindings</span></span><br><span data-ttu-id="73e8e-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-699">
         - TextCoercion</span></span><br><span data-ttu-id="73e8e-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-701">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="73e8e-701">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="73e8e-702">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="73e8e-703">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="73e8e-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="73e8e-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="73e8e-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="73e8e-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="73e8e-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-707">- BindingEvents</span></span><br><span data-ttu-id="73e8e-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-708">
         - CompressedFile</span></span><br><span data-ttu-id="73e8e-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="73e8e-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="73e8e-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-710">
         - DocumentEvents</span></span><br><span data-ttu-id="73e8e-711">
         - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-711">
         - File</span></span><br><span data-ttu-id="73e8e-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="73e8e-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-713">
         - MatrixBindings</span></span><br><span data-ttu-id="73e8e-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="73e8e-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="73e8e-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-716">
         - PdfFile</span></span><br><span data-ttu-id="73e8e-717">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-717">
         - Selection</span></span><br><span data-ttu-id="73e8e-718">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-718">
         - Settings</span></span><br><span data-ttu-id="73e8e-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-719">
         - TableBindings</span></span><br><span data-ttu-id="73e8e-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-720">
         - TableCoercion</span></span><br><span data-ttu-id="73e8e-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="73e8e-721">
         - TextBindings</span></span><br><span data-ttu-id="73e8e-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-722">
         - TextCoercion</span></span><br><span data-ttu-id="73e8e-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="73e8e-724">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="73e8e-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="73e8e-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="73e8e-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="73e8e-726">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="73e8e-726">Platform</span></span></th>
    <th><span data-ttu-id="73e8e-727">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="73e8e-727">Extension points</span></span></th>
    <th><span data-ttu-id="73e8e-728">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="73e8e-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="73e8e-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="73e8e-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-730">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="73e8e-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="73e8e-731">- Contenu</span><span class="sxs-lookup"><span data-stu-id="73e8e-731">- Content</span></span><br><span data-ttu-id="73e8e-732">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-732">
         - TaskPane</span></span><br><span data-ttu-id="73e8e-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="73e8e-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="73e8e-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="73e8e-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="73e8e-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="73e8e-737">- ActiveView</span></span><br><span data-ttu-id="73e8e-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-738">
         - CompressedFile</span></span><br><span data-ttu-id="73e8e-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-739">
         - DocumentEvents</span></span><br><span data-ttu-id="73e8e-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-740">
         - File</span></span><br><span data-ttu-id="73e8e-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-741">
         - PdfFile</span></span><br><span data-ttu-id="73e8e-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-742">
         - Selection</span></span><br><span data-ttu-id="73e8e-743">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-743">
         - Settings</span></span><br><span data-ttu-id="73e8e-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-745">Office pour Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-745">Office on Windows</span></span><br><span data-ttu-id="73e8e-746">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="73e8e-746">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="73e8e-747">- Contenu</span><span class="sxs-lookup"><span data-stu-id="73e8e-747">- Content</span></span><br><span data-ttu-id="73e8e-748">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-748">
         - TaskPane</span></span><br><span data-ttu-id="73e8e-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="73e8e-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="73e8e-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="73e8e-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="73e8e-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="73e8e-753">- ActiveView</span></span><br><span data-ttu-id="73e8e-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-754">
         - CompressedFile</span></span><br><span data-ttu-id="73e8e-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-755">
         - DocumentEvents</span></span><br><span data-ttu-id="73e8e-756">
         - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-756">
         - File</span></span><br><span data-ttu-id="73e8e-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-757">
         - PdfFile</span></span><br><span data-ttu-id="73e8e-758">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-758">
         - Selection</span></span><br><span data-ttu-id="73e8e-759">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-759">
         - Settings</span></span><br><span data-ttu-id="73e8e-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-761">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-761">Office 2019 on Windows</span></span><br><span data-ttu-id="73e8e-762">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="73e8e-763">- Contenu</span><span class="sxs-lookup"><span data-stu-id="73e8e-763">- Content</span></span><br><span data-ttu-id="73e8e-764">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-764">
         - TaskPane</span></span><br><span data-ttu-id="73e8e-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="73e8e-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="73e8e-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="73e8e-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="73e8e-768">- ActiveView</span></span><br><span data-ttu-id="73e8e-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-769">
         - CompressedFile</span></span><br><span data-ttu-id="73e8e-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-770">
         - DocumentEvents</span></span><br><span data-ttu-id="73e8e-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-771">
         - File</span></span><br><span data-ttu-id="73e8e-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-772">
         - PdfFile</span></span><br><span data-ttu-id="73e8e-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-773">
         - Selection</span></span><br><span data-ttu-id="73e8e-774">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-774">
         - Settings</span></span><br><span data-ttu-id="73e8e-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-776">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-776">Office 2016 on Windows</span></span><br><span data-ttu-id="73e8e-777">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="73e8e-778">- Contenu</span><span class="sxs-lookup"><span data-stu-id="73e8e-778">- Content</span></span><br><span data-ttu-id="73e8e-779">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="73e8e-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="73e8e-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="73e8e-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="73e8e-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="73e8e-782">- ActiveView</span></span><br><span data-ttu-id="73e8e-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-783">
         - CompressedFile</span></span><br><span data-ttu-id="73e8e-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-784">
         - DocumentEvents</span></span><br><span data-ttu-id="73e8e-785">
         - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-785">
         - File</span></span><br><span data-ttu-id="73e8e-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-786">
         - PdfFile</span></span><br><span data-ttu-id="73e8e-787">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-787">
         - Selection</span></span><br><span data-ttu-id="73e8e-788">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-788">
         - Settings</span></span><br><span data-ttu-id="73e8e-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-790">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-790">Office 2013 on Windows</span></span><br><span data-ttu-id="73e8e-791">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="73e8e-792">- Contenu</span><span class="sxs-lookup"><span data-stu-id="73e8e-792">- Content</span></span><br><span data-ttu-id="73e8e-793">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="73e8e-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="73e8e-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="73e8e-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="73e8e-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="73e8e-796">- ActiveView</span></span><br><span data-ttu-id="73e8e-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-797">
         - CompressedFile</span></span><br><span data-ttu-id="73e8e-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-798">
         - DocumentEvents</span></span><br><span data-ttu-id="73e8e-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-799">
         - File</span></span><br><span data-ttu-id="73e8e-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-800">
         - PdfFile</span></span><br><span data-ttu-id="73e8e-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-801">
         - Selection</span></span><br><span data-ttu-id="73e8e-802">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-802">
         - Settings</span></span><br><span data-ttu-id="73e8e-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-804">Office sur iPad</span><span class="sxs-lookup"><span data-stu-id="73e8e-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="73e8e-805">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="73e8e-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="73e8e-806">- Contenu</span><span class="sxs-lookup"><span data-stu-id="73e8e-806">- Content</span></span><br><span data-ttu-id="73e8e-807">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="73e8e-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="73e8e-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="73e8e-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="73e8e-810">- ActiveView</span></span><br><span data-ttu-id="73e8e-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-811">
         - CompressedFile</span></span><br><span data-ttu-id="73e8e-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-812">
         - DocumentEvents</span></span><br><span data-ttu-id="73e8e-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-813">
         - File</span></span><br><span data-ttu-id="73e8e-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-814">
         - PdfFile</span></span><br><span data-ttu-id="73e8e-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-815">
         - Selection</span></span><br><span data-ttu-id="73e8e-816">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-816">
         - Settings</span></span><br><span data-ttu-id="73e8e-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-818">Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="73e8e-818">Office apps on Mac</span></span><br><span data-ttu-id="73e8e-819">(connecté à l’abonnement Office 365)</span><span class="sxs-lookup"><span data-stu-id="73e8e-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="73e8e-820">- Contenu</span><span class="sxs-lookup"><span data-stu-id="73e8e-820">- Content</span></span><br><span data-ttu-id="73e8e-821">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-821">
         - TaskPane</span></span><br><span data-ttu-id="73e8e-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="73e8e-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="73e8e-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="73e8e-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="73e8e-826">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="73e8e-826">- ActiveView</span></span><br><span data-ttu-id="73e8e-827">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-827">
         - CompressedFile</span></span><br><span data-ttu-id="73e8e-828">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-828">
         - DocumentEvents</span></span><br><span data-ttu-id="73e8e-829">
         - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-829">
         - File</span></span><br><span data-ttu-id="73e8e-830">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-830">
         - PdfFile</span></span><br><span data-ttu-id="73e8e-831">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-831">
         - Selection</span></span><br><span data-ttu-id="73e8e-832">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-832">
         - Settings</span></span><br><span data-ttu-id="73e8e-833">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-833">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-834">Office 2019 sur Mac</span><span class="sxs-lookup"><span data-stu-id="73e8e-834">Office 2019 for Mac</span></span><br><span data-ttu-id="73e8e-835">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-835">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="73e8e-836">- Contenu</span><span class="sxs-lookup"><span data-stu-id="73e8e-836">- Content</span></span><br><span data-ttu-id="73e8e-837">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-837">
         - TaskPane</span></span><br><span data-ttu-id="73e8e-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="73e8e-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="73e8e-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="73e8e-841">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="73e8e-841">- ActiveView</span></span><br><span data-ttu-id="73e8e-842">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-842">
         - CompressedFile</span></span><br><span data-ttu-id="73e8e-843">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-843">
         - DocumentEvents</span></span><br><span data-ttu-id="73e8e-844">
         - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-844">
         - File</span></span><br><span data-ttu-id="73e8e-845">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-845">
         - PdfFile</span></span><br><span data-ttu-id="73e8e-846">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-846">
         - Selection</span></span><br><span data-ttu-id="73e8e-847">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-847">
         - Settings</span></span><br><span data-ttu-id="73e8e-848">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-848">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-849">Office 2016 sur Mac</span><span class="sxs-lookup"><span data-stu-id="73e8e-849">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="73e8e-850">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-850">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="73e8e-851">- Contenu</span><span class="sxs-lookup"><span data-stu-id="73e8e-851">- Content</span></span><br><span data-ttu-id="73e8e-852">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-852">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="73e8e-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="73e8e-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="73e8e-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="73e8e-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="73e8e-855">- ActiveView</span></span><br><span data-ttu-id="73e8e-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-856">
         - CompressedFile</span></span><br><span data-ttu-id="73e8e-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-857">
         - DocumentEvents</span></span><br><span data-ttu-id="73e8e-858">
         - File</span><span class="sxs-lookup"><span data-stu-id="73e8e-858">
         - File</span></span><br><span data-ttu-id="73e8e-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="73e8e-859">
         - PdfFile</span></span><br><span data-ttu-id="73e8e-860">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-860">
         - Selection</span></span><br><span data-ttu-id="73e8e-861">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-861">
         - Settings</span></span><br><span data-ttu-id="73e8e-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-862">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="73e8e-863">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="73e8e-863">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="73e8e-864">OneNote</span><span class="sxs-lookup"><span data-stu-id="73e8e-864">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="73e8e-865">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="73e8e-865">Platform</span></span></th>
    <th><span data-ttu-id="73e8e-866">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="73e8e-866">Extension points</span></span></th>
    <th><span data-ttu-id="73e8e-867">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="73e8e-867">API requirement sets</span></span></th>
    <th><span data-ttu-id="73e8e-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="73e8e-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-869">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="73e8e-869">Office on the web</span></span></td>
    <td> <span data-ttu-id="73e8e-870">- Contenu</span><span class="sxs-lookup"><span data-stu-id="73e8e-870">- Content</span></span><br><span data-ttu-id="73e8e-871">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-871">
         - TaskPane</span></span><br><span data-ttu-id="73e8e-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="73e8e-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="73e8e-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="73e8e-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="73e8e-876">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="73e8e-876">- DocumentEvents</span></span><br><span data-ttu-id="73e8e-877">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-877">
         - HtmlCoercion</span></span><br><span data-ttu-id="73e8e-878">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="73e8e-878">
         - Settings</span></span><br><span data-ttu-id="73e8e-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-879">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="73e8e-880">Projet</span><span class="sxs-lookup"><span data-stu-id="73e8e-880">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="73e8e-881">Plateforme</span><span class="sxs-lookup"><span data-stu-id="73e8e-881">Platform</span></span></th>
    <th><span data-ttu-id="73e8e-882">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="73e8e-882">Extension points</span></span></th>
    <th><span data-ttu-id="73e8e-883">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="73e8e-883">API requirement sets</span></span></th>
    <th><span data-ttu-id="73e8e-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="73e8e-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-885">Office 2019 sur Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-885">Office 2019 on Windows</span></span><br><span data-ttu-id="73e8e-886">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-886">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="73e8e-887">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-887">- TaskPane</span></span></td>
    <td> <span data-ttu-id="73e8e-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="73e8e-889">- Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-889">- Selection</span></span><br><span data-ttu-id="73e8e-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-890">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-891">Office 2016 sur Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-891">Office 2016 on Windows</span></span><br><span data-ttu-id="73e8e-892">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="73e8e-893">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="73e8e-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="73e8e-895">- Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-895">- Selection</span></span><br><span data-ttu-id="73e8e-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="73e8e-897">Office 2013 sur Windows</span><span class="sxs-lookup"><span data-stu-id="73e8e-897">Office 2013 on Windows</span></span><br><span data-ttu-id="73e8e-898">(achat définitif)</span><span class="sxs-lookup"><span data-stu-id="73e8e-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="73e8e-899">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="73e8e-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="73e8e-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="73e8e-901">- Selection</span><span class="sxs-lookup"><span data-stu-id="73e8e-901">- Selection</span></span><br><span data-ttu-id="73e8e-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="73e8e-902">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="73e8e-903">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="73e8e-903">See also</span></span>

- [<span data-ttu-id="73e8e-904">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-904">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="73e8e-905">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="73e8e-905">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="73e8e-906">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="73e8e-906">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="73e8e-907">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="73e8e-907">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="73e8e-908">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="73e8e-908">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="73e8e-909">Historique des mises à jour d’Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="73e8e-909">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="73e8e-910">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="73e8e-910">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="73e8e-911">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="73e8e-911">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="73e8e-912">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="73e8e-912">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="73e8e-913">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="73e8e-913">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="73e8e-914">Historique des mises à jour d’Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="73e8e-914">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
