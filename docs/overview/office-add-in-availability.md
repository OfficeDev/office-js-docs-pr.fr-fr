---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint, OneNote et Project.
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 28a6d0e4c86d05855ed9d24461dbeb77454d2b48
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872129"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="9fc21-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="9fc21-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="9fc21-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="9fc21-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="9fc21-p102">Le numéro de build pour Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que les ensembles de conditions requises des API communes, ExcelApi 1.1 et WordApi 1.1.</span><span class="sxs-lookup"><span data-stu-id="9fc21-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>
>
> <span data-ttu-id="9fc21-108">Le numéro de build d’un achat définitif d’Office 2019 est 16.0.10827.20150.</span><span class="sxs-lookup"><span data-stu-id="9fc21-108">The build number for a one-time purchase of Office 2019 is 16.0.10827.20150.</span></span>

## <a name="excel"></a><span data-ttu-id="9fc21-109">Excel</span><span class="sxs-lookup"><span data-stu-id="9fc21-109">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="9fc21-110">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="9fc21-110">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="9fc21-111">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="9fc21-111">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="9fc21-112">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="9fc21-112">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="9fc21-113"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="9fc21-113"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-114">Office Online</span><span class="sxs-lookup"><span data-stu-id="9fc21-114">Office Online</span></span></td>
    <td> <span data-ttu-id="9fc21-115">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-115">- TaskPane</span></span><br><span data-ttu-id="9fc21-116">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="9fc21-116">
        - Content</span></span><br><span data-ttu-id="9fc21-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="9fc21-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="9fc21-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9fc21-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9fc21-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9fc21-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9fc21-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9fc21-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9fc21-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9fc21-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9fc21-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9fc21-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-127">
        - BindingEvents</span></span><br><span data-ttu-id="9fc21-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-128">
        - CompressedFile</span></span><br><span data-ttu-id="9fc21-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-129">
        - DocumentEvents</span></span><br><span data-ttu-id="9fc21-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-130">
        - File</span></span><br><span data-ttu-id="9fc21-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-131">
        - MatrixBindings</span></span><br><span data-ttu-id="9fc21-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-132">
        - MatrixCoercion</span></span><br><span data-ttu-id="9fc21-133">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-133">
        - Selection</span></span><br><span data-ttu-id="9fc21-134">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-134">
        - Settings</span></span><br><span data-ttu-id="9fc21-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-135">
        - TableBindings</span></span><br><span data-ttu-id="9fc21-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-136">
        - TableCoercion</span></span><br><span data-ttu-id="9fc21-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-137">
        - TextBindings</span></span><br><span data-ttu-id="9fc21-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-138">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-139">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="9fc21-139">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="9fc21-140">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-140">- TaskPane</span></span><br><span data-ttu-id="9fc21-141">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="9fc21-141">
        - Content</span></span><br><span data-ttu-id="9fc21-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="9fc21-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="9fc21-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9fc21-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9fc21-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9fc21-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9fc21-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9fc21-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9fc21-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9fc21-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9fc21-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9fc21-152">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-152">
        - BindingEvents</span></span><br><span data-ttu-id="9fc21-153">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-153">
        - CompressedFile</span></span><br><span data-ttu-id="9fc21-154">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-154">
        - DocumentEvents</span></span><br><span data-ttu-id="9fc21-155">
        - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-155">
        - File</span></span><br><span data-ttu-id="9fc21-156">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-156">
        - MatrixBindings</span></span><br><span data-ttu-id="9fc21-157">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-157">
        - MatrixCoercion</span></span><br><span data-ttu-id="9fc21-158">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-158">
        - Selection</span></span><br><span data-ttu-id="9fc21-159">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-159">
        - Settings</span></span><br><span data-ttu-id="9fc21-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-160">
        - TableBindings</span></span><br><span data-ttu-id="9fc21-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-161">
        - TableCoercion</span></span><br><span data-ttu-id="9fc21-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-162">
        - TextBindings</span></span><br><span data-ttu-id="9fc21-163">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-163">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-164">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="9fc21-164">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="9fc21-165">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-165">- TaskPane</span></span><br><span data-ttu-id="9fc21-166">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="9fc21-166">
        - Content</span></span><br><span data-ttu-id="9fc21-167">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-167">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="9fc21-168">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-168">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9fc21-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9fc21-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9fc21-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9fc21-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9fc21-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9fc21-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9fc21-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9fc21-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9fc21-177">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-177">- BindingEvents</span></span><br><span data-ttu-id="9fc21-178">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-178">
        - CompressedFile</span></span><br><span data-ttu-id="9fc21-179">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-179">
        - DocumentEvents</span></span><br><span data-ttu-id="9fc21-180">
        - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-180">
        - File</span></span><br><span data-ttu-id="9fc21-181">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-181">
        - ImageCoercion</span></span><br><span data-ttu-id="9fc21-182">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-182">
        - MatrixBindings</span></span><br><span data-ttu-id="9fc21-183">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-183">
        - MatrixCoercion</span></span><br><span data-ttu-id="9fc21-184">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-184">
        - Selection</span></span><br><span data-ttu-id="9fc21-185">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-185">
        - Settings</span></span><br><span data-ttu-id="9fc21-186">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-186">
        - TableBindings</span></span><br><span data-ttu-id="9fc21-187">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-187">
        - TableCoercion</span></span><br><span data-ttu-id="9fc21-188">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-188">
        - TextBindings</span></span><br><span data-ttu-id="9fc21-189">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-189">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-190">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="9fc21-190">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="9fc21-191">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-191">- TaskPane</span></span><br><span data-ttu-id="9fc21-192">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="9fc21-192">
        - Content</span></span></td>
    <td><span data-ttu-id="9fc21-193">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-193">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9fc21-194">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="9fc21-194">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="9fc21-195">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-195">- BindingEvents</span></span><br><span data-ttu-id="9fc21-196">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-196">
        - CompressedFile</span></span><br><span data-ttu-id="9fc21-197">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-197">
        - DocumentEvents</span></span><br><span data-ttu-id="9fc21-198">
        - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-198">
        - File</span></span><br><span data-ttu-id="9fc21-199">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-199">
        - ImageCoercion</span></span><br><span data-ttu-id="9fc21-200">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-200">
        - MatrixBindings</span></span><br><span data-ttu-id="9fc21-201">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-201">
        - MatrixCoercion</span></span><br><span data-ttu-id="9fc21-202">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-202">
        - Selection</span></span><br><span data-ttu-id="9fc21-203">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-203">
        - Settings</span></span><br><span data-ttu-id="9fc21-204">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-204">
        - TableBindings</span></span><br><span data-ttu-id="9fc21-205">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-205">
        - TableCoercion</span></span><br><span data-ttu-id="9fc21-206">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-206">
        - TextBindings</span></span><br><span data-ttu-id="9fc21-207">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-207">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-208">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="9fc21-208">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="9fc21-209">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-209">
        - TaskPane</span></span><br><span data-ttu-id="9fc21-210">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="9fc21-210">
        - Content</span></span></td>
    <td>  <span data-ttu-id="9fc21-211">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9fc21-211">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="9fc21-212">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-212">
        - BindingEvents</span></span><br><span data-ttu-id="9fc21-213">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-213">
        - CompressedFile</span></span><br><span data-ttu-id="9fc21-214">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-214">
        - DocumentEvents</span></span><br><span data-ttu-id="9fc21-215">
        - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-215">
        - File</span></span><br><span data-ttu-id="9fc21-216">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-216">
        - ImageCoercion</span></span><br><span data-ttu-id="9fc21-217">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-217">
        - MatrixBindings</span></span><br><span data-ttu-id="9fc21-218">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-218">
        - MatrixCoercion</span></span><br><span data-ttu-id="9fc21-219">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-219">
        - Selection</span></span><br><span data-ttu-id="9fc21-220">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-220">
        - Settings</span></span><br><span data-ttu-id="9fc21-221">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-221">
        - TableBindings</span></span><br><span data-ttu-id="9fc21-222">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-222">
        - TableCoercion</span></span><br><span data-ttu-id="9fc21-223">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-223">
        - TextBindings</span></span><br><span data-ttu-id="9fc21-224">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-224">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-225">Office 365 pour iPad</span><span class="sxs-lookup"><span data-stu-id="9fc21-225">Office 365 for iPad</span></span></td>
    <td><span data-ttu-id="9fc21-226">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-226">- TaskPane</span></span><br><span data-ttu-id="9fc21-227">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="9fc21-227">
        - Content</span></span></td>
    <td><span data-ttu-id="9fc21-228">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-228">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9fc21-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9fc21-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9fc21-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9fc21-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9fc21-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9fc21-234">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-234">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9fc21-235">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-235">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9fc21-236">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-236">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9fc21-237">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-237">- BindingEvents</span></span><br><span data-ttu-id="9fc21-238">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-238">
        - CompressedFile</span></span><br><span data-ttu-id="9fc21-239">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-239">
        - DocumentEvents</span></span><br><span data-ttu-id="9fc21-240">
        - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-240">
        - File</span></span><br><span data-ttu-id="9fc21-241">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-241">
        - ImageCoercion</span></span><br><span data-ttu-id="9fc21-242">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-242">
        - MatrixBindings</span></span><br><span data-ttu-id="9fc21-243">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-243">
        - MatrixCoercion</span></span><br><span data-ttu-id="9fc21-244">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-244">
        - Selection</span></span><br><span data-ttu-id="9fc21-245">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-245">
        - Settings</span></span><br><span data-ttu-id="9fc21-246">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-246">
        - TableBindings</span></span><br><span data-ttu-id="9fc21-247">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-247">
        - TableCoercion</span></span><br><span data-ttu-id="9fc21-248">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-248">
        - TextBindings</span></span><br><span data-ttu-id="9fc21-249">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-249">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-250">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="9fc21-250">Office 365 for Mac</span></span></td>
    <td><span data-ttu-id="9fc21-251">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-251">- TaskPane</span></span><br><span data-ttu-id="9fc21-252">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="9fc21-252">
        - Content</span></span><br><span data-ttu-id="9fc21-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="9fc21-254">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-254">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9fc21-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9fc21-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9fc21-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9fc21-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9fc21-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9fc21-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9fc21-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9fc21-262">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-262">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9fc21-263">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-263">- BindingEvents</span></span><br><span data-ttu-id="9fc21-264">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-264">
        - CompressedFile</span></span><br><span data-ttu-id="9fc21-265">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-265">
        - DocumentEvents</span></span><br><span data-ttu-id="9fc21-266">
        - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-266">
        - File</span></span><br><span data-ttu-id="9fc21-267">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-267">
        - ImageCoercion</span></span><br><span data-ttu-id="9fc21-268">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-268">
        - MatrixBindings</span></span><br><span data-ttu-id="9fc21-269">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-269">
        - MatrixCoercion</span></span><br><span data-ttu-id="9fc21-270">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-270">
        - PdfFile</span></span><br><span data-ttu-id="9fc21-271">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-271">
        - Selection</span></span><br><span data-ttu-id="9fc21-272">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-272">
        - Settings</span></span><br><span data-ttu-id="9fc21-273">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-273">
        - TableBindings</span></span><br><span data-ttu-id="9fc21-274">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-274">
        - TableCoercion</span></span><br><span data-ttu-id="9fc21-275">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-275">
        - TextBindings</span></span><br><span data-ttu-id="9fc21-276">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-276">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-277">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="9fc21-277">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="9fc21-278">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-278">- TaskPane</span></span><br><span data-ttu-id="9fc21-279">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="9fc21-279">
        - Content</span></span><br><span data-ttu-id="9fc21-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="9fc21-281">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-281">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9fc21-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9fc21-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9fc21-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9fc21-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9fc21-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9fc21-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9fc21-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9fc21-289">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-289">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9fc21-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-290">- BindingEvents</span></span><br><span data-ttu-id="9fc21-291">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-291">
        - CompressedFile</span></span><br><span data-ttu-id="9fc21-292">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-292">
        - DocumentEvents</span></span><br><span data-ttu-id="9fc21-293">
        - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-293">
        - File</span></span><br><span data-ttu-id="9fc21-294">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-294">
        - ImageCoercion</span></span><br><span data-ttu-id="9fc21-295">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-295">
        - MatrixBindings</span></span><br><span data-ttu-id="9fc21-296">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-296">
        - MatrixCoercion</span></span><br><span data-ttu-id="9fc21-297">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-297">
        - PdfFile</span></span><br><span data-ttu-id="9fc21-298">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-298">
        - Selection</span></span><br><span data-ttu-id="9fc21-299">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-299">
        - Settings</span></span><br><span data-ttu-id="9fc21-300">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-300">
        - TableBindings</span></span><br><span data-ttu-id="9fc21-301">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-301">
        - TableCoercion</span></span><br><span data-ttu-id="9fc21-302">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-302">
        - TextBindings</span></span><br><span data-ttu-id="9fc21-303">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-303">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-304">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="9fc21-304">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="9fc21-305">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-305">- TaskPane</span></span><br><span data-ttu-id="9fc21-306">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="9fc21-306">
        - Content</span></span></td>
    <td><span data-ttu-id="9fc21-307">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-307">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9fc21-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="9fc21-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="9fc21-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-309">- BindingEvents</span></span><br><span data-ttu-id="9fc21-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-310">
        - CompressedFile</span></span><br><span data-ttu-id="9fc21-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-311">
        - DocumentEvents</span></span><br><span data-ttu-id="9fc21-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-312">
        - File</span></span><br><span data-ttu-id="9fc21-313">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-313">
        - ImageCoercion</span></span><br><span data-ttu-id="9fc21-314">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-314">
        - MatrixBindings</span></span><br><span data-ttu-id="9fc21-315">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-315">
        - MatrixCoercion</span></span><br><span data-ttu-id="9fc21-316">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-316">
        - PdfFile</span></span><br><span data-ttu-id="9fc21-317">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-317">
        - Selection</span></span><br><span data-ttu-id="9fc21-318">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-318">
        - Settings</span></span><br><span data-ttu-id="9fc21-319">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-319">
        - TableBindings</span></span><br><span data-ttu-id="9fc21-320">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-320">
        - TableCoercion</span></span><br><span data-ttu-id="9fc21-321">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-321">
        - TextBindings</span></span><br><span data-ttu-id="9fc21-322">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-322">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="9fc21-323">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="9fc21-323">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="9fc21-324">Outlook</span><span class="sxs-lookup"><span data-stu-id="9fc21-324">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9fc21-325">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="9fc21-325">Platform</span></span></th>
    <th><span data-ttu-id="9fc21-326">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="9fc21-326">Extension points</span></span></th>
    <th><span data-ttu-id="9fc21-327">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="9fc21-327">API requirement sets</span></span></th>
    <th><span data-ttu-id="9fc21-328"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="9fc21-328"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-329">Office Online</span><span class="sxs-lookup"><span data-stu-id="9fc21-329">Office Online</span></span></td>
    <td> <span data-ttu-id="9fc21-330">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="9fc21-330">- Mail Read</span></span><br><span data-ttu-id="9fc21-331">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="9fc21-331">
      - Mail Compose</span></span><br><span data-ttu-id="9fc21-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9fc21-333">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-333">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9fc21-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9fc21-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9fc21-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9fc21-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9fc21-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="9fc21-339">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-339">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="9fc21-340">Non disponible</span><span class="sxs-lookup"><span data-stu-id="9fc21-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-341">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="9fc21-341">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="9fc21-342">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="9fc21-342">- Mail Read</span></span><br><span data-ttu-id="9fc21-343">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="9fc21-343">
      - Mail Compose</span></span><br><span data-ttu-id="9fc21-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="9fc21-345">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="9fc21-345">
      - Modules</span></span></td>
    <td> <span data-ttu-id="9fc21-346">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-346">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9fc21-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9fc21-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9fc21-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9fc21-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9fc21-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="9fc21-352">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-352">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="9fc21-353">Non disponible</span><span class="sxs-lookup"><span data-stu-id="9fc21-353">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-354">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="9fc21-354">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="9fc21-355">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="9fc21-355">- Mail Read</span></span><br><span data-ttu-id="9fc21-356">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="9fc21-356">
      - Mail Compose</span></span><br><span data-ttu-id="9fc21-357">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-357">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="9fc21-358">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="9fc21-358">
      - Modules</span></span></td>
    <td> <span data-ttu-id="9fc21-359">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-359">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9fc21-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9fc21-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9fc21-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9fc21-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9fc21-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="9fc21-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="9fc21-366">Non disponible</span><span class="sxs-lookup"><span data-stu-id="9fc21-366">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-367">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="9fc21-367">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="9fc21-368">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="9fc21-368">- Mail Read</span></span><br><span data-ttu-id="9fc21-369">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="9fc21-369">
      - Mail Compose</span></span><br><span data-ttu-id="9fc21-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="9fc21-371">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="9fc21-371">
      - Modules</span></span></td>
    <td> <span data-ttu-id="9fc21-372">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-372">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9fc21-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9fc21-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9fc21-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="9fc21-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="9fc21-376">Non disponible</span><span class="sxs-lookup"><span data-stu-id="9fc21-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-377">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="9fc21-377">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="9fc21-378">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="9fc21-378">- Mail Read</span></span><br><span data-ttu-id="9fc21-379">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="9fc21-379">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="9fc21-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9fc21-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9fc21-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="9fc21-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="9fc21-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="9fc21-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="9fc21-384">Non disponible</span><span class="sxs-lookup"><span data-stu-id="9fc21-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-385">Office 365 pour iOS</span><span class="sxs-lookup"><span data-stu-id="9fc21-385">Office 365 for iOS</span></span></td>
    <td> <span data-ttu-id="9fc21-386">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="9fc21-386">- Mail Read</span></span><br><span data-ttu-id="9fc21-387">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-387">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9fc21-388">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-388">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9fc21-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9fc21-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9fc21-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9fc21-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="9fc21-393">Non disponible</span><span class="sxs-lookup"><span data-stu-id="9fc21-393">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-394">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="9fc21-394">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="9fc21-395">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="9fc21-395">- Mail Read</span></span><br><span data-ttu-id="9fc21-396">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="9fc21-396">
      - Mail Compose</span></span><br><span data-ttu-id="9fc21-397">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-397">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9fc21-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9fc21-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9fc21-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9fc21-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9fc21-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9fc21-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="9fc21-404">Non disponible</span><span class="sxs-lookup"><span data-stu-id="9fc21-404">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-405">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="9fc21-405">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="9fc21-406">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="9fc21-406">- Mail Read</span></span><br><span data-ttu-id="9fc21-407">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="9fc21-407">
      - Mail Compose</span></span><br><span data-ttu-id="9fc21-408">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-408">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9fc21-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9fc21-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9fc21-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9fc21-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9fc21-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9fc21-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="9fc21-415">Non disponible</span><span class="sxs-lookup"><span data-stu-id="9fc21-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-416">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="9fc21-416">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="9fc21-417">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="9fc21-417">- Mail Read</span></span><br><span data-ttu-id="9fc21-418">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="9fc21-418">
      - Mail Compose</span></span><br><span data-ttu-id="9fc21-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9fc21-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9fc21-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9fc21-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9fc21-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9fc21-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9fc21-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="9fc21-426">Non disponible</span><span class="sxs-lookup"><span data-stu-id="9fc21-426">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-427">Office 365 pour Android</span><span class="sxs-lookup"><span data-stu-id="9fc21-427">Office 365 for Android</span></span></td>
    <td> <span data-ttu-id="9fc21-428">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="9fc21-428">- Mail Read</span></span><br><span data-ttu-id="9fc21-429">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-429">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9fc21-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9fc21-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9fc21-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9fc21-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9fc21-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="9fc21-435">Non disponible</span><span class="sxs-lookup"><span data-stu-id="9fc21-435">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="9fc21-436">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="9fc21-436">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="9fc21-437">Word</span><span class="sxs-lookup"><span data-stu-id="9fc21-437">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9fc21-438">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="9fc21-438">Platform</span></span></th>
    <th><span data-ttu-id="9fc21-439">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="9fc21-439">Extension points</span></span></th>
    <th><span data-ttu-id="9fc21-440">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="9fc21-440">API requirement sets</span></span></th>
    <th><span data-ttu-id="9fc21-441"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="9fc21-441"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-442">Office Online</span><span class="sxs-lookup"><span data-stu-id="9fc21-442">Office Online</span></span></td>
    <td> <span data-ttu-id="9fc21-443">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-443">- TaskPane</span></span><br><span data-ttu-id="9fc21-444">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-444">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9fc21-445">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-445">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9fc21-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9fc21-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9fc21-448">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-448">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9fc21-449">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-449">- BindingEvents</span></span><br><span data-ttu-id="9fc21-450">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9fc21-450">
         - CustomXmlParts</span></span><br><span data-ttu-id="9fc21-451">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-451">
         - DocumentEvents</span></span><br><span data-ttu-id="9fc21-452">
         - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-452">
         - File</span></span><br><span data-ttu-id="9fc21-453">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-453">
         - HtmlCoercion</span></span><br><span data-ttu-id="9fc21-454">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-454">
         - ImageCoercion</span></span><br><span data-ttu-id="9fc21-455">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-455">
         - MatrixBindings</span></span><br><span data-ttu-id="9fc21-456">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-456">
         - MatrixCoercion</span></span><br><span data-ttu-id="9fc21-457">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-457">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9fc21-458">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-458">
         - PdfFile</span></span><br><span data-ttu-id="9fc21-459">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-459">
         - Selection</span></span><br><span data-ttu-id="9fc21-460">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-460">
         - Settings</span></span><br><span data-ttu-id="9fc21-461">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-461">
         - TableBindings</span></span><br><span data-ttu-id="9fc21-462">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-462">
         - TableCoercion</span></span><br><span data-ttu-id="9fc21-463">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-463">
         - TextBindings</span></span><br><span data-ttu-id="9fc21-464">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-464">
         - TextCoercion</span></span><br><span data-ttu-id="9fc21-465">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-465">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-466">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="9fc21-466">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="9fc21-467">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-467">- TaskPane</span></span><br><span data-ttu-id="9fc21-468">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-468">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9fc21-469">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-469">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9fc21-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9fc21-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9fc21-472">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-472">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9fc21-473">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-473">- BindingEvents</span></span><br><span data-ttu-id="9fc21-474">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-474">
         - CompressedFile</span></span><br><span data-ttu-id="9fc21-475">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9fc21-475">
         - CustomXmlParts</span></span><br><span data-ttu-id="9fc21-476">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-476">
         - DocumentEvents</span></span><br><span data-ttu-id="9fc21-477">
         - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-477">
         - File</span></span><br><span data-ttu-id="9fc21-478">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-478">
         - HtmlCoercion</span></span><br><span data-ttu-id="9fc21-479">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-479">
         - ImageCoercion</span></span><br><span data-ttu-id="9fc21-480">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-480">
         - MatrixBindings</span></span><br><span data-ttu-id="9fc21-481">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-481">
         - MatrixCoercion</span></span><br><span data-ttu-id="9fc21-482">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-482">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9fc21-483">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-483">
         - PdfFile</span></span><br><span data-ttu-id="9fc21-484">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-484">
         - Selection</span></span><br><span data-ttu-id="9fc21-485">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-485">
         - Settings</span></span><br><span data-ttu-id="9fc21-486">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-486">
         - TableBindings</span></span><br><span data-ttu-id="9fc21-487">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-487">
         - TableCoercion</span></span><br><span data-ttu-id="9fc21-488">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-488">
         - TextBindings</span></span><br><span data-ttu-id="9fc21-489">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-489">
         - TextCoercion</span></span><br><span data-ttu-id="9fc21-490">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-490">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-491">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="9fc21-491">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="9fc21-492">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-492">- TaskPane</span></span><br><span data-ttu-id="9fc21-493">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-493">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9fc21-494">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-494">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9fc21-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9fc21-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9fc21-497">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-497">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9fc21-498">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-498">- BindingEvents</span></span><br><span data-ttu-id="9fc21-499">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-499">
         - CompressedFile</span></span><br><span data-ttu-id="9fc21-500">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9fc21-500">
         - CustomXmlParts</span></span><br><span data-ttu-id="9fc21-501">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-501">
         - DocumentEvents</span></span><br><span data-ttu-id="9fc21-502">
         - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-502">
         - File</span></span><br><span data-ttu-id="9fc21-503">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-503">
         - HtmlCoercion</span></span><br><span data-ttu-id="9fc21-504">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-504">
         - ImageCoercion</span></span><br><span data-ttu-id="9fc21-505">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-505">
         - MatrixBindings</span></span><br><span data-ttu-id="9fc21-506">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-506">
         - MatrixCoercion</span></span><br><span data-ttu-id="9fc21-507">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-507">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9fc21-508">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-508">
         - PdfFile</span></span><br><span data-ttu-id="9fc21-509">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-509">
         - Selection</span></span><br><span data-ttu-id="9fc21-510">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-510">
         - Settings</span></span><br><span data-ttu-id="9fc21-511">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-511">
         - TableBindings</span></span><br><span data-ttu-id="9fc21-512">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-512">
         - TableCoercion</span></span><br><span data-ttu-id="9fc21-513">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-513">
         - TextBindings</span></span><br><span data-ttu-id="9fc21-514">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-514">
         - TextCoercion</span></span><br><span data-ttu-id="9fc21-515">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-515">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-516">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="9fc21-516">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="9fc21-517">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-517">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9fc21-518">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-518">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9fc21-519">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="9fc21-519">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="9fc21-520">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-520">- BindingEvents</span></span><br><span data-ttu-id="9fc21-521">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-521">
         - CompressedFile</span></span><br><span data-ttu-id="9fc21-522">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9fc21-522">
         - CustomXmlParts</span></span><br><span data-ttu-id="9fc21-523">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-523">
         - DocumentEvents</span></span><br><span data-ttu-id="9fc21-524">
         - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-524">
         - File</span></span><br><span data-ttu-id="9fc21-525">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-525">
         - HtmlCoercion</span></span><br><span data-ttu-id="9fc21-526">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-526">
         - ImageCoercion</span></span><br><span data-ttu-id="9fc21-527">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-527">
         - MatrixBindings</span></span><br><span data-ttu-id="9fc21-528">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-528">
         - MatrixCoercion</span></span><br><span data-ttu-id="9fc21-529">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-529">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9fc21-530">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-530">
         - PdfFile</span></span><br><span data-ttu-id="9fc21-531">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-531">
         - Selection</span></span><br><span data-ttu-id="9fc21-532">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-532">
         - Settings</span></span><br><span data-ttu-id="9fc21-533">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-533">
         - TableBindings</span></span><br><span data-ttu-id="9fc21-534">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-534">
         - TableCoercion</span></span><br><span data-ttu-id="9fc21-535">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-535">
         - TextBindings</span></span><br><span data-ttu-id="9fc21-536">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-536">
         - TextCoercion</span></span><br><span data-ttu-id="9fc21-537">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-537">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-538">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="9fc21-538">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="9fc21-539">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-539">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9fc21-540">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9fc21-540">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="9fc21-541">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-541">- BindingEvents</span></span><br><span data-ttu-id="9fc21-542">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-542">
         - CompressedFile</span></span><br><span data-ttu-id="9fc21-543">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9fc21-543">
         - CustomXmlParts</span></span><br><span data-ttu-id="9fc21-544">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-544">
         - DocumentEvents</span></span><br><span data-ttu-id="9fc21-545">
         - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-545">
         - File</span></span><br><span data-ttu-id="9fc21-546">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-546">
         - HtmlCoercion</span></span><br><span data-ttu-id="9fc21-547">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-547">
         - ImageCoercion</span></span><br><span data-ttu-id="9fc21-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-548">
         - MatrixBindings</span></span><br><span data-ttu-id="9fc21-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="9fc21-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9fc21-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-551">
         - PdfFile</span></span><br><span data-ttu-id="9fc21-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-552">
         - Selection</span></span><br><span data-ttu-id="9fc21-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-553">
         - Settings</span></span><br><span data-ttu-id="9fc21-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-554">
         - TableBindings</span></span><br><span data-ttu-id="9fc21-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-555">
         - TableCoercion</span></span><br><span data-ttu-id="9fc21-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-556">
         - TextBindings</span></span><br><span data-ttu-id="9fc21-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-557">
         - TextCoercion</span></span><br><span data-ttu-id="9fc21-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-558">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-559">Office 365 pour iPad</span><span class="sxs-lookup"><span data-stu-id="9fc21-559">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="9fc21-560">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-560">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9fc21-561">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-561">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9fc21-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9fc21-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9fc21-564">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="9fc21-564">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="9fc21-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-565">- BindingEvents</span></span><br><span data-ttu-id="9fc21-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-566">
         - CompressedFile</span></span><br><span data-ttu-id="9fc21-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9fc21-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="9fc21-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-568">
         - DocumentEvents</span></span><br><span data-ttu-id="9fc21-569">
         - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-569">
         - File</span></span><br><span data-ttu-id="9fc21-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="9fc21-571">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-571">
         - ImageCoercion</span></span><br><span data-ttu-id="9fc21-572">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-572">
         - MatrixBindings</span></span><br><span data-ttu-id="9fc21-573">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-573">
         - MatrixCoercion</span></span><br><span data-ttu-id="9fc21-574">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-574">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9fc21-575">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-575">
         - PdfFile</span></span><br><span data-ttu-id="9fc21-576">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-576">
         - Selection</span></span><br><span data-ttu-id="9fc21-577">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-577">
         - Settings</span></span><br><span data-ttu-id="9fc21-578">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-578">
         - TableBindings</span></span><br><span data-ttu-id="9fc21-579">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-579">
         - TableCoercion</span></span><br><span data-ttu-id="9fc21-580">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-580">
         - TextBindings</span></span><br><span data-ttu-id="9fc21-581">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-581">
         - TextCoercion</span></span><br><span data-ttu-id="9fc21-582">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-582">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-583">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="9fc21-583">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="9fc21-584">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-584">- TaskPane</span></span><br><span data-ttu-id="9fc21-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9fc21-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9fc21-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9fc21-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9fc21-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="9fc21-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="9fc21-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-590">- BindingEvents</span></span><br><span data-ttu-id="9fc21-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-591">
         - CompressedFile</span></span><br><span data-ttu-id="9fc21-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9fc21-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="9fc21-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-593">
         - DocumentEvents</span></span><br><span data-ttu-id="9fc21-594">
         - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-594">
         - File</span></span><br><span data-ttu-id="9fc21-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="9fc21-596">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-596">
         - ImageCoercion</span></span><br><span data-ttu-id="9fc21-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-597">
         - MatrixBindings</span></span><br><span data-ttu-id="9fc21-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="9fc21-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9fc21-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-600">
         - PdfFile</span></span><br><span data-ttu-id="9fc21-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-601">
         - Selection</span></span><br><span data-ttu-id="9fc21-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-602">
         - Settings</span></span><br><span data-ttu-id="9fc21-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-603">
         - TableBindings</span></span><br><span data-ttu-id="9fc21-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-604">
         - TableCoercion</span></span><br><span data-ttu-id="9fc21-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-605">
         - TextBindings</span></span><br><span data-ttu-id="9fc21-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-606">
         - TextCoercion</span></span><br><span data-ttu-id="9fc21-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-608">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="9fc21-608">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="9fc21-609">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-609">- TaskPane</span></span><br><span data-ttu-id="9fc21-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9fc21-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9fc21-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9fc21-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9fc21-614">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="9fc21-614">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="9fc21-615">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-615">- BindingEvents</span></span><br><span data-ttu-id="9fc21-616">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-616">
         - CompressedFile</span></span><br><span data-ttu-id="9fc21-617">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9fc21-617">
         - CustomXmlParts</span></span><br><span data-ttu-id="9fc21-618">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-618">
         - DocumentEvents</span></span><br><span data-ttu-id="9fc21-619">
         - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-619">
         - File</span></span><br><span data-ttu-id="9fc21-620">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-620">
         - HtmlCoercion</span></span><br><span data-ttu-id="9fc21-621">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-621">
         - ImageCoercion</span></span><br><span data-ttu-id="9fc21-622">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-622">
         - MatrixBindings</span></span><br><span data-ttu-id="9fc21-623">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-623">
         - MatrixCoercion</span></span><br><span data-ttu-id="9fc21-624">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-624">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9fc21-625">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-625">
         - PdfFile</span></span><br><span data-ttu-id="9fc21-626">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-626">
         - Selection</span></span><br><span data-ttu-id="9fc21-627">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-627">
         - Settings</span></span><br><span data-ttu-id="9fc21-628">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-628">
         - TableBindings</span></span><br><span data-ttu-id="9fc21-629">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-629">
         - TableCoercion</span></span><br><span data-ttu-id="9fc21-630">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-630">
         - TextBindings</span></span><br><span data-ttu-id="9fc21-631">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-631">
         - TextCoercion</span></span><br><span data-ttu-id="9fc21-632">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-632">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-633">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="9fc21-633">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="9fc21-634">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-634">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9fc21-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9fc21-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="9fc21-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="9fc21-637">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-637">- BindingEvents</span></span><br><span data-ttu-id="9fc21-638">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-638">
         - CompressedFile</span></span><br><span data-ttu-id="9fc21-639">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9fc21-639">
         - CustomXmlParts</span></span><br><span data-ttu-id="9fc21-640">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-640">
         - DocumentEvents</span></span><br><span data-ttu-id="9fc21-641">
         - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-641">
         - File</span></span><br><span data-ttu-id="9fc21-642">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-642">
         - HtmlCoercion</span></span><br><span data-ttu-id="9fc21-643">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-643">
         - ImageCoercion</span></span><br><span data-ttu-id="9fc21-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-644">
         - MatrixBindings</span></span><br><span data-ttu-id="9fc21-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="9fc21-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9fc21-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-647">
         - PdfFile</span></span><br><span data-ttu-id="9fc21-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-648">
         - Selection</span></span><br><span data-ttu-id="9fc21-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-649">
         - Settings</span></span><br><span data-ttu-id="9fc21-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-650">
         - TableBindings</span></span><br><span data-ttu-id="9fc21-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-651">
         - TableCoercion</span></span><br><span data-ttu-id="9fc21-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9fc21-652">
         - TextBindings</span></span><br><span data-ttu-id="9fc21-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-653">
         - TextCoercion</span></span><br><span data-ttu-id="9fc21-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-654">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="9fc21-655">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="9fc21-655">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="9fc21-656">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="9fc21-656">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9fc21-657">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="9fc21-657">Platform</span></span></th>
    <th><span data-ttu-id="9fc21-658">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="9fc21-658">Extension points</span></span></th>
    <th><span data-ttu-id="9fc21-659">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="9fc21-659">API requirement sets</span></span></th>
    <th><span data-ttu-id="9fc21-660"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="9fc21-660"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-661">Office Online</span><span class="sxs-lookup"><span data-stu-id="9fc21-661">Office Online</span></span></td>
    <td> <span data-ttu-id="9fc21-662">- Contenu</span><span class="sxs-lookup"><span data-stu-id="9fc21-662">- Content</span></span><br><span data-ttu-id="9fc21-663">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-663">
         - TaskPane</span></span><br><span data-ttu-id="9fc21-664">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-664">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9fc21-665">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-665">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9fc21-666">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9fc21-666">- ActiveView</span></span><br><span data-ttu-id="9fc21-667">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-667">
         - CompressedFile</span></span><br><span data-ttu-id="9fc21-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-668">
         - DocumentEvents</span></span><br><span data-ttu-id="9fc21-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-669">
         - File</span></span><br><span data-ttu-id="9fc21-670">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-670">
         - ImageCoercion</span></span><br><span data-ttu-id="9fc21-671">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-671">
         - PdfFile</span></span><br><span data-ttu-id="9fc21-672">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-672">
         - Selection</span></span><br><span data-ttu-id="9fc21-673">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-673">
         - Settings</span></span><br><span data-ttu-id="9fc21-674">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-674">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-675">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="9fc21-675">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="9fc21-676">- Contenu</span><span class="sxs-lookup"><span data-stu-id="9fc21-676">- Content</span></span><br><span data-ttu-id="9fc21-677">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-677">
         - TaskPane</span></span><br><span data-ttu-id="9fc21-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9fc21-679">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-679">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9fc21-680">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9fc21-680">- ActiveView</span></span><br><span data-ttu-id="9fc21-681">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-681">
         - CompressedFile</span></span><br><span data-ttu-id="9fc21-682">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-682">
         - DocumentEvents</span></span><br><span data-ttu-id="9fc21-683">
         - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-683">
         - File</span></span><br><span data-ttu-id="9fc21-684">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-684">
         - ImageCoercion</span></span><br><span data-ttu-id="9fc21-685">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-685">
         - PdfFile</span></span><br><span data-ttu-id="9fc21-686">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-686">
         - Selection</span></span><br><span data-ttu-id="9fc21-687">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-687">
         - Settings</span></span><br><span data-ttu-id="9fc21-688">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-688">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-689">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="9fc21-689">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="9fc21-690">- Contenu</span><span class="sxs-lookup"><span data-stu-id="9fc21-690">- Content</span></span><br><span data-ttu-id="9fc21-691">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-691">
         - TaskPane</span></span><br><span data-ttu-id="9fc21-692">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-692">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9fc21-693">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-693">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9fc21-694">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9fc21-694">- ActiveView</span></span><br><span data-ttu-id="9fc21-695">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-695">
         - CompressedFile</span></span><br><span data-ttu-id="9fc21-696">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-696">
         - DocumentEvents</span></span><br><span data-ttu-id="9fc21-697">
         - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-697">
         - File</span></span><br><span data-ttu-id="9fc21-698">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-698">
         - ImageCoercion</span></span><br><span data-ttu-id="9fc21-699">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-699">
         - PdfFile</span></span><br><span data-ttu-id="9fc21-700">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-700">
         - Selection</span></span><br><span data-ttu-id="9fc21-701">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-701">
         - Settings</span></span><br><span data-ttu-id="9fc21-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-702">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-703">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="9fc21-703">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="9fc21-704">- Contenu</span><span class="sxs-lookup"><span data-stu-id="9fc21-704">- Content</span></span><br><span data-ttu-id="9fc21-705">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-705">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="9fc21-706">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9fc21-706">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="9fc21-707">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9fc21-707">- ActiveView</span></span><br><span data-ttu-id="9fc21-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-708">
         - CompressedFile</span></span><br><span data-ttu-id="9fc21-709">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-709">
         - DocumentEvents</span></span><br><span data-ttu-id="9fc21-710">
         - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-710">
         - File</span></span><br><span data-ttu-id="9fc21-711">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-711">
         - ImageCoercion</span></span><br><span data-ttu-id="9fc21-712">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-712">
         - PdfFile</span></span><br><span data-ttu-id="9fc21-713">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-713">
         - Selection</span></span><br><span data-ttu-id="9fc21-714">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-714">
         - Settings</span></span><br><span data-ttu-id="9fc21-715">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-715">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-716">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="9fc21-716">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="9fc21-717">- Contenu</span><span class="sxs-lookup"><span data-stu-id="9fc21-717">- Content</span></span><br><span data-ttu-id="9fc21-718">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-718">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="9fc21-719">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9fc21-719">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="9fc21-720">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9fc21-720">- ActiveView</span></span><br><span data-ttu-id="9fc21-721">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-721">
         - CompressedFile</span></span><br><span data-ttu-id="9fc21-722">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-722">
         - DocumentEvents</span></span><br><span data-ttu-id="9fc21-723">
         - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-723">
         - File</span></span><br><span data-ttu-id="9fc21-724">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-724">
         - ImageCoercion</span></span><br><span data-ttu-id="9fc21-725">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-725">
         - PdfFile</span></span><br><span data-ttu-id="9fc21-726">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-726">
         - Selection</span></span><br><span data-ttu-id="9fc21-727">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-727">
         - Settings</span></span><br><span data-ttu-id="9fc21-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-728">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-729">Office 365 pour iPad</span><span class="sxs-lookup"><span data-stu-id="9fc21-729">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="9fc21-730">- Contenu</span><span class="sxs-lookup"><span data-stu-id="9fc21-730">- Content</span></span><br><span data-ttu-id="9fc21-731">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-731">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="9fc21-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="9fc21-733">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9fc21-733">- ActiveView</span></span><br><span data-ttu-id="9fc21-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-734">
         - CompressedFile</span></span><br><span data-ttu-id="9fc21-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-735">
         - DocumentEvents</span></span><br><span data-ttu-id="9fc21-736">
         - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-736">
         - File</span></span><br><span data-ttu-id="9fc21-737">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-737">
         - PdfFile</span></span><br><span data-ttu-id="9fc21-738">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-738">
         - Selection</span></span><br><span data-ttu-id="9fc21-739">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-739">
         - Settings</span></span><br><span data-ttu-id="9fc21-740">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-740">
         - TextCoercion</span></span><br><span data-ttu-id="9fc21-741">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-741">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-742">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="9fc21-742">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="9fc21-743">- Contenu</span><span class="sxs-lookup"><span data-stu-id="9fc21-743">- Content</span></span><br><span data-ttu-id="9fc21-744">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-744">
         - TaskPane</span></span><br><span data-ttu-id="9fc21-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9fc21-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9fc21-747">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9fc21-747">- ActiveView</span></span><br><span data-ttu-id="9fc21-748">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-748">
         - CompressedFile</span></span><br><span data-ttu-id="9fc21-749">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-749">
         - DocumentEvents</span></span><br><span data-ttu-id="9fc21-750">
         - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-750">
         - File</span></span><br><span data-ttu-id="9fc21-751">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-751">
         - ImageCoercion</span></span><br><span data-ttu-id="9fc21-752">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-752">
         - PdfFile</span></span><br><span data-ttu-id="9fc21-753">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-753">
         - Selection</span></span><br><span data-ttu-id="9fc21-754">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-754">
         - Settings</span></span><br><span data-ttu-id="9fc21-755">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-755">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-756">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="9fc21-756">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="9fc21-757">- Contenu</span><span class="sxs-lookup"><span data-stu-id="9fc21-757">- Content</span></span><br><span data-ttu-id="9fc21-758">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-758">
         - TaskPane</span></span><br><span data-ttu-id="9fc21-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9fc21-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9fc21-761">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9fc21-761">- ActiveView</span></span><br><span data-ttu-id="9fc21-762">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-762">
         - CompressedFile</span></span><br><span data-ttu-id="9fc21-763">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-763">
         - DocumentEvents</span></span><br><span data-ttu-id="9fc21-764">
         - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-764">
         - File</span></span><br><span data-ttu-id="9fc21-765">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-765">
         - ImageCoercion</span></span><br><span data-ttu-id="9fc21-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-766">
         - PdfFile</span></span><br><span data-ttu-id="9fc21-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-767">
         - Selection</span></span><br><span data-ttu-id="9fc21-768">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-768">
         - Settings</span></span><br><span data-ttu-id="9fc21-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-770">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="9fc21-770">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="9fc21-771">- Contenu</span><span class="sxs-lookup"><span data-stu-id="9fc21-771">- Content</span></span><br><span data-ttu-id="9fc21-772">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-772">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="9fc21-773">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9fc21-773">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="9fc21-774">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9fc21-774">- ActiveView</span></span><br><span data-ttu-id="9fc21-775">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-775">
         - CompressedFile</span></span><br><span data-ttu-id="9fc21-776">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-776">
         - DocumentEvents</span></span><br><span data-ttu-id="9fc21-777">
         - File</span><span class="sxs-lookup"><span data-stu-id="9fc21-777">
         - File</span></span><br><span data-ttu-id="9fc21-778">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-778">
         - ImageCoercion</span></span><br><span data-ttu-id="9fc21-779">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9fc21-779">
         - PdfFile</span></span><br><span data-ttu-id="9fc21-780">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-780">
         - Selection</span></span><br><span data-ttu-id="9fc21-781">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-781">
         - Settings</span></span><br><span data-ttu-id="9fc21-782">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-782">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="9fc21-783">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="9fc21-783">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="9fc21-784">OneNote</span><span class="sxs-lookup"><span data-stu-id="9fc21-784">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9fc21-785">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="9fc21-785">Platform</span></span></th>
    <th><span data-ttu-id="9fc21-786">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="9fc21-786">Extension points</span></span></th>
    <th><span data-ttu-id="9fc21-787">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="9fc21-787">API requirement sets</span></span></th>
    <th><span data-ttu-id="9fc21-788"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="9fc21-788"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-789">Office Online</span><span class="sxs-lookup"><span data-stu-id="9fc21-789">Office Online</span></span></td>
    <td> <span data-ttu-id="9fc21-790">- Contenu</span><span class="sxs-lookup"><span data-stu-id="9fc21-790">- Content</span></span><br><span data-ttu-id="9fc21-791">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-791">
         - TaskPane</span></span><br><span data-ttu-id="9fc21-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9fc21-793">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-793">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="9fc21-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9fc21-795">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9fc21-795">- DocumentEvents</span></span><br><span data-ttu-id="9fc21-796">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-796">
         - HtmlCoercion</span></span><br><span data-ttu-id="9fc21-797">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-797">
         - ImageCoercion</span></span><br><span data-ttu-id="9fc21-798">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9fc21-798">
         - Settings</span></span><br><span data-ttu-id="9fc21-799">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-799">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="9fc21-800">Projet</span><span class="sxs-lookup"><span data-stu-id="9fc21-800">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9fc21-801">Plateforme</span><span class="sxs-lookup"><span data-stu-id="9fc21-801">Platform</span></span></th>
    <th><span data-ttu-id="9fc21-802">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="9fc21-802">Extension points</span></span></th>
    <th><span data-ttu-id="9fc21-803">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="9fc21-803">API requirement sets</span></span></th>
    <th><span data-ttu-id="9fc21-804"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="9fc21-804"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-805">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="9fc21-805">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="9fc21-806">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-806">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9fc21-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9fc21-808">- Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-808">- Selection</span></span><br><span data-ttu-id="9fc21-809">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-809">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-810">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="9fc21-810">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="9fc21-811">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-811">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9fc21-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9fc21-813">- Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-813">- Selection</span></span><br><span data-ttu-id="9fc21-814">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-814">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9fc21-815">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="9fc21-815">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="9fc21-816">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-816">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9fc21-817">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9fc21-817">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9fc21-818">- Selection</span><span class="sxs-lookup"><span data-stu-id="9fc21-818">- Selection</span></span><br><span data-ttu-id="9fc21-819">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9fc21-819">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="9fc21-820">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="9fc21-820">See also</span></span>

- [<span data-ttu-id="9fc21-821">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-821">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="9fc21-822">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="9fc21-822">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="9fc21-823">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="9fc21-823">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="9fc21-824">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="9fc21-824">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
