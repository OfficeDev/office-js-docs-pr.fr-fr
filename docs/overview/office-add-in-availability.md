---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint, OneNote et Project.
ms.date: 11/07/2018
ms.openlocfilehash: c3da40be21c0e569028dd10e93e33760ba2bd39d
ms.sourcegitcommit: 3e84d616e69f39eeeeea773f2431e7d674c4a9f5
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/22/2018
ms.locfileid: "26644752"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="672c7-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="672c7-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="672c7-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span><span class="sxs-lookup"><span data-stu-id="672c7-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="672c7-p102">Le numéro de build pour Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que les ensembles de conditions requises des API communes, ExcelApi 1.1 et WordApi 1.1.</span><span class="sxs-lookup"><span data-stu-id="672c7-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="672c7-108">Excel</span><span class="sxs-lookup"><span data-stu-id="672c7-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="672c7-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="672c7-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="672c7-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="672c7-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="672c7-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="672c7-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="672c7-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="672c7-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="672c7-113">Office Online</span></span></td>
    <td> <span data-ttu-id="672c7-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-114">- TaskPane</span></span><br><span data-ttu-id="672c7-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="672c7-115">
        - Content</span></span><br><span data-ttu-id="672c7-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="672c7-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="672c7-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="672c7-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="672c7-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="672c7-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="672c7-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="672c7-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="672c7-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="672c7-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="672c7-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="672c7-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="672c7-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="672c7-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="672c7-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="672c7-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="672c7-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-126">
        - BindingEvents</span></span><br><span data-ttu-id="672c7-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-127">
        - CompressedFile</span></span><br><span data-ttu-id="672c7-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-128">
        - DocumentEvents</span></span><br><span data-ttu-id="672c7-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="672c7-129">
        - File</span></span><br><span data-ttu-id="672c7-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-130">
        - MatrixBindings</span></span><br><span data-ttu-id="672c7-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="672c7-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-132">
        - Selection</span></span><br><span data-ttu-id="672c7-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-133">
        - Settings</span></span><br><span data-ttu-id="672c7-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-134">
        - TableBindings</span></span><br><span data-ttu-id="672c7-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-135">
        - TableCoercion</span></span><br><span data-ttu-id="672c7-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-136">
        - TextBindings</span></span><br><span data-ttu-id="672c7-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-138">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="672c7-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="672c7-139">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-139">
        - TaskPane</span></span><br><span data-ttu-id="672c7-140">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="672c7-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="672c7-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="672c7-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-142">
        - BindingEvents</span></span><br><span data-ttu-id="672c7-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-143">
        - CompressedFile</span></span><br><span data-ttu-id="672c7-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-144">
        - DocumentEvents</span></span><br><span data-ttu-id="672c7-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="672c7-145">
        - File</span></span><br><span data-ttu-id="672c7-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-146">
        - ImageCoercion</span></span><br><span data-ttu-id="672c7-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-147">
        - MatrixBindings</span></span><br><span data-ttu-id="672c7-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="672c7-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-149">
        - Selection</span></span><br><span data-ttu-id="672c7-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-150">
        - Settings</span></span><br><span data-ttu-id="672c7-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-151">
        - TableBindings</span></span><br><span data-ttu-id="672c7-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-152">
        - TableCoercion</span></span><br><span data-ttu-id="672c7-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-153">
        - TextBindings</span></span><br><span data-ttu-id="672c7-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-155">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="672c7-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="672c7-156">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-156">- TaskPane</span></span><br><span data-ttu-id="672c7-157">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="672c7-157">
        - Content</span></span><br><span data-ttu-id="672c7-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="672c7-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="672c7-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="672c7-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="672c7-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="672c7-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="672c7-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="672c7-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="672c7-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="672c7-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="672c7-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="672c7-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="672c7-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="672c7-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="672c7-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="672c7-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-168">- BindingEvents</span></span><br><span data-ttu-id="672c7-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-169">
        - CompressedFile</span></span><br><span data-ttu-id="672c7-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-170">
        - DocumentEvents</span></span><br><span data-ttu-id="672c7-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="672c7-171">
        - File</span></span><br><span data-ttu-id="672c7-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-172">
        - ImageCoercion</span></span><br><span data-ttu-id="672c7-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-173">
        - MatrixBindings</span></span><br><span data-ttu-id="672c7-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-174">
        - MatrixCoercion</span></span><br><span data-ttu-id="672c7-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-175">
        - Selection</span></span><br><span data-ttu-id="672c7-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-176">
        - Settings</span></span><br><span data-ttu-id="672c7-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-177">
        - TableBindings</span></span><br><span data-ttu-id="672c7-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-178">
        - TableCoercion</span></span><br><span data-ttu-id="672c7-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-179">
        - TextBindings</span></span><br><span data-ttu-id="672c7-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-180">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-181">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="672c7-181">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="672c7-182">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-182">- TaskPane</span></span><br><span data-ttu-id="672c7-183">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="672c7-183">
        - Content</span></span><br><span data-ttu-id="672c7-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="672c7-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="672c7-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="672c7-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="672c7-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="672c7-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="672c7-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="672c7-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="672c7-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="672c7-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="672c7-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="672c7-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="672c7-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="672c7-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="672c7-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="672c7-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-194">- BindingEvents</span></span><br><span data-ttu-id="672c7-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-195">
        - CompressedFile</span></span><br><span data-ttu-id="672c7-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-196">
        - DocumentEvents</span></span><br><span data-ttu-id="672c7-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="672c7-197">
        - File</span></span><br><span data-ttu-id="672c7-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-198">
        - ImageCoercion</span></span><br><span data-ttu-id="672c7-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-199">
        - MatrixBindings</span></span><br><span data-ttu-id="672c7-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="672c7-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-201">
        - Selection</span></span><br><span data-ttu-id="672c7-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-202">
        - Settings</span></span><br><span data-ttu-id="672c7-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-203">
        - TableBindings</span></span><br><span data-ttu-id="672c7-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-204">
        - TableCoercion</span></span><br><span data-ttu-id="672c7-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-205">
        - TextBindings</span></span><br><span data-ttu-id="672c7-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-207">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="672c7-207">Office for iPad</span></span></td>
    <td><span data-ttu-id="672c7-208">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-208">- TaskPane</span></span><br><span data-ttu-id="672c7-209">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="672c7-209">
        - Content</span></span></td>
    <td><span data-ttu-id="672c7-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="672c7-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="672c7-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="672c7-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="672c7-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="672c7-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="672c7-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="672c7-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="672c7-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="672c7-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="672c7-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="672c7-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="672c7-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="672c7-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="672c7-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-219">- BindingEvents</span></span><br><span data-ttu-id="672c7-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-220">
        - CompressedFile</span></span><br><span data-ttu-id="672c7-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-221">
        - DocumentEvents</span></span><br><span data-ttu-id="672c7-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="672c7-222">
        - File</span></span><br><span data-ttu-id="672c7-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-223">
        - ImageCoercion</span></span><br><span data-ttu-id="672c7-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-224">
        - MatrixBindings</span></span><br><span data-ttu-id="672c7-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="672c7-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-226">
        - Selection</span></span><br><span data-ttu-id="672c7-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-227">
        - Settings</span></span><br><span data-ttu-id="672c7-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-228">
        - TableBindings</span></span><br><span data-ttu-id="672c7-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-229">
        - TableCoercion</span></span><br><span data-ttu-id="672c7-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-230">
        - TextBindings</span></span><br><span data-ttu-id="672c7-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-232">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="672c7-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="672c7-233">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-233">- TaskPane</span></span><br><span data-ttu-id="672c7-234">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="672c7-234">
        - Content</span></span><br><span data-ttu-id="672c7-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="672c7-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="672c7-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="672c7-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="672c7-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="672c7-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="672c7-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="672c7-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="672c7-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="672c7-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="672c7-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="672c7-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="672c7-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="672c7-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="672c7-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="672c7-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-245">- BindingEvents</span></span><br><span data-ttu-id="672c7-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-246">
        - CompressedFile</span></span><br><span data-ttu-id="672c7-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-247">
        - DocumentEvents</span></span><br><span data-ttu-id="672c7-248">
        - File</span><span class="sxs-lookup"><span data-stu-id="672c7-248">
        - File</span></span><br><span data-ttu-id="672c7-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-249">
        - ImageCoercion</span></span><br><span data-ttu-id="672c7-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-250">
        - MatrixBindings</span></span><br><span data-ttu-id="672c7-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-251">
        - MatrixCoercion</span></span><br><span data-ttu-id="672c7-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="672c7-252">
        - PdfFile</span></span><br><span data-ttu-id="672c7-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-253">
        - Selection</span></span><br><span data-ttu-id="672c7-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-254">
        - Settings</span></span><br><span data-ttu-id="672c7-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-255">
        - TableBindings</span></span><br><span data-ttu-id="672c7-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-256">
        - TableCoercion</span></span><br><span data-ttu-id="672c7-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-257">
        - TextBindings</span></span><br><span data-ttu-id="672c7-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-259">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="672c7-259">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="672c7-260">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-260">- TaskPane</span></span><br><span data-ttu-id="672c7-261">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="672c7-261">
        - Content</span></span><br><span data-ttu-id="672c7-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="672c7-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="672c7-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="672c7-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="672c7-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="672c7-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="672c7-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="672c7-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="672c7-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="672c7-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="672c7-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="672c7-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="672c7-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="672c7-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="672c7-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="672c7-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-272">- BindingEvents</span></span><br><span data-ttu-id="672c7-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-273">
        - CompressedFile</span></span><br><span data-ttu-id="672c7-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-274">
        - DocumentEvents</span></span><br><span data-ttu-id="672c7-275">
        - File</span><span class="sxs-lookup"><span data-stu-id="672c7-275">
        - File</span></span><br><span data-ttu-id="672c7-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-276">
        - ImageCoercion</span></span><br><span data-ttu-id="672c7-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-277">
        - MatrixBindings</span></span><br><span data-ttu-id="672c7-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-278">
        - MatrixCoercion</span></span><br><span data-ttu-id="672c7-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="672c7-279">
        - PdfFile</span></span><br><span data-ttu-id="672c7-280">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-280">
        - Selection</span></span><br><span data-ttu-id="672c7-281">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-281">
        - Settings</span></span><br><span data-ttu-id="672c7-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-282">
        - TableBindings</span></span><br><span data-ttu-id="672c7-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-283">
        - TableCoercion</span></span><br><span data-ttu-id="672c7-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-284">
        - TextBindings</span></span><br><span data-ttu-id="672c7-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-285">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="672c7-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="672c7-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="672c7-287">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="672c7-287">Platform</span></span></th>
    <th><span data-ttu-id="672c7-288">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="672c7-288">Extension points</span></span></th>
    <th><span data-ttu-id="672c7-289">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="672c7-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="672c7-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="672c7-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="672c7-291">Office Online</span></span></td>
    <td> <span data-ttu-id="672c7-292">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="672c7-292">- Mail Read</span></span><br><span data-ttu-id="672c7-293">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="672c7-293">
      - Mail Compose</span></span><br><span data-ttu-id="672c7-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="672c7-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="672c7-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="672c7-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="672c7-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="672c7-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="672c7-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="672c7-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="672c7-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="672c7-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="672c7-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="672c7-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="672c7-302">Non disponible</span><span class="sxs-lookup"><span data-stu-id="672c7-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-303">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="672c7-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="672c7-304">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="672c7-304">- Mail Read</span></span><br><span data-ttu-id="672c7-305">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="672c7-305">
      - Mail Compose</span></span><br><span data-ttu-id="672c7-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="672c7-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="672c7-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="672c7-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="672c7-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="672c7-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="672c7-311">Non disponible</span><span class="sxs-lookup"><span data-stu-id="672c7-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-312">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="672c7-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="672c7-313">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="672c7-313">- Mail Read</span></span><br><span data-ttu-id="672c7-314">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="672c7-314">
      - Mail Compose</span></span><br><span data-ttu-id="672c7-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="672c7-316">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="672c7-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="672c7-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="672c7-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="672c7-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="672c7-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="672c7-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="672c7-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="672c7-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="672c7-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="672c7-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="672c7-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="672c7-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="672c7-324">Non disponible</span><span class="sxs-lookup"><span data-stu-id="672c7-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-325">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="672c7-325">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="672c7-326">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="672c7-326">- Mail Read</span></span><br><span data-ttu-id="672c7-327">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="672c7-327">
      - Mail Compose</span></span><br><span data-ttu-id="672c7-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="672c7-329">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="672c7-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="672c7-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="672c7-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="672c7-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="672c7-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="672c7-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="672c7-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="672c7-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="672c7-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="672c7-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="672c7-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="672c7-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="672c7-337">Non disponible</span><span class="sxs-lookup"><span data-stu-id="672c7-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-338">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="672c7-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="672c7-339">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="672c7-339">- Mail Read</span></span><br><span data-ttu-id="672c7-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="672c7-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="672c7-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="672c7-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="672c7-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="672c7-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="672c7-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="672c7-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="672c7-346">Non disponible</span><span class="sxs-lookup"><span data-stu-id="672c7-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-347">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="672c7-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="672c7-348">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="672c7-348">- Mail Read</span></span><br><span data-ttu-id="672c7-349">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="672c7-349">
      - Mail Compose</span></span><br><span data-ttu-id="672c7-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="672c7-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="672c7-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="672c7-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="672c7-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="672c7-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="672c7-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="672c7-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="672c7-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="672c7-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="672c7-357">Non disponible</span><span class="sxs-lookup"><span data-stu-id="672c7-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-358">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="672c7-358">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="672c7-359">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="672c7-359">- Mail Read</span></span><br><span data-ttu-id="672c7-360">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="672c7-360">
      - Mail Compose</span></span><br><span data-ttu-id="672c7-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="672c7-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="672c7-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="672c7-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="672c7-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="672c7-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="672c7-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="672c7-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="672c7-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="672c7-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="672c7-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="672c7-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="672c7-369">Non disponible</span><span class="sxs-lookup"><span data-stu-id="672c7-369">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-370">Office pour Android</span><span class="sxs-lookup"><span data-stu-id="672c7-370">Office for Android</span></span></td>
    <td> <span data-ttu-id="672c7-371">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="672c7-371">- Mail Read</span></span><br><span data-ttu-id="672c7-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="672c7-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="672c7-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="672c7-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="672c7-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="672c7-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="672c7-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="672c7-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="672c7-378">Non disponible</span><span class="sxs-lookup"><span data-stu-id="672c7-378">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="672c7-379">Word</span><span class="sxs-lookup"><span data-stu-id="672c7-379">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="672c7-380">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="672c7-380">Platform</span></span></th>
    <th><span data-ttu-id="672c7-381">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="672c7-381">Extension points</span></span></th>
    <th><span data-ttu-id="672c7-382">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="672c7-382">API requirement sets</span></span></th>
    <th><span data-ttu-id="672c7-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="672c7-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-384">Office Online</span><span class="sxs-lookup"><span data-stu-id="672c7-384">Office Online</span></span></td>
    <td> <span data-ttu-id="672c7-385">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-385">- TaskPane</span></span><br><span data-ttu-id="672c7-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="672c7-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="672c7-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="672c7-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="672c7-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="672c7-391">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-391">- BindingEvents</span></span><br><span data-ttu-id="672c7-392">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="672c7-392">
         - CustomXmlParts</span></span><br><span data-ttu-id="672c7-393">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-393">
         - DocumentEvents</span></span><br><span data-ttu-id="672c7-394">
         - File</span><span class="sxs-lookup"><span data-stu-id="672c7-394">
         - File</span></span><br><span data-ttu-id="672c7-395">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-395">
         - HtmlCoercion</span></span><br><span data-ttu-id="672c7-396">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-396">
         - ImageCoercion</span></span><br><span data-ttu-id="672c7-397">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-397">
         - MatrixBindings</span></span><br><span data-ttu-id="672c7-398">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-398">
         - MatrixCoercion</span></span><br><span data-ttu-id="672c7-399">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-399">
         - OoxmlCoercion</span></span><br><span data-ttu-id="672c7-400">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="672c7-400">
         - PdfFile</span></span><br><span data-ttu-id="672c7-401">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-401">
         - Selection</span></span><br><span data-ttu-id="672c7-402">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-402">
         - Settings</span></span><br><span data-ttu-id="672c7-403">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-403">
         - TableBindings</span></span><br><span data-ttu-id="672c7-404">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-404">
         - TableCoercion</span></span><br><span data-ttu-id="672c7-405">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-405">
         - TextBindings</span></span><br><span data-ttu-id="672c7-406">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-406">
         - TextCoercion</span></span><br><span data-ttu-id="672c7-407">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="672c7-407">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-408">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="672c7-408">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="672c7-409">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-409">- TaskPane</span></span></td>
    <td> <span data-ttu-id="672c7-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="672c7-411">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-411">- BindingEvents</span></span><br><span data-ttu-id="672c7-412">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-412">
         - CompressedFile</span></span><br><span data-ttu-id="672c7-413">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="672c7-413">
         - CustomXmlParts</span></span><br><span data-ttu-id="672c7-414">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-414">
         - DocumentEvents</span></span><br><span data-ttu-id="672c7-415">
         - File</span><span class="sxs-lookup"><span data-stu-id="672c7-415">
         - File</span></span><br><span data-ttu-id="672c7-416">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-416">
         - HtmlCoercion</span></span><br><span data-ttu-id="672c7-417">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-417">
         - ImageCoercion</span></span><br><span data-ttu-id="672c7-418">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-418">
         - MatrixBindings</span></span><br><span data-ttu-id="672c7-419">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-419">
         - MatrixCoercion</span></span><br><span data-ttu-id="672c7-420">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-420">
         - OoxmlCoercion</span></span><br><span data-ttu-id="672c7-421">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="672c7-421">
         - PdfFile</span></span><br><span data-ttu-id="672c7-422">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-422">
         - Selection</span></span><br><span data-ttu-id="672c7-423">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-423">
         - Settings</span></span><br><span data-ttu-id="672c7-424">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-424">
         - TableBindings</span></span><br><span data-ttu-id="672c7-425">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-425">
         - TableCoercion</span></span><br><span data-ttu-id="672c7-426">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-426">
         - TextBindings</span></span><br><span data-ttu-id="672c7-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-427">
         - TextCoercion</span></span><br><span data-ttu-id="672c7-428">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="672c7-428">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-429">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="672c7-429">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="672c7-430">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-430">- TaskPane</span></span><br><span data-ttu-id="672c7-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="672c7-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="672c7-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="672c7-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="672c7-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="672c7-436">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-436">- BindingEvents</span></span><br><span data-ttu-id="672c7-437">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-437">
         - CompressedFile</span></span><br><span data-ttu-id="672c7-438">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="672c7-438">
         - CustomXmlParts</span></span><br><span data-ttu-id="672c7-439">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-439">
         - DocumentEvents</span></span><br><span data-ttu-id="672c7-440">
         - File</span><span class="sxs-lookup"><span data-stu-id="672c7-440">
         - File</span></span><br><span data-ttu-id="672c7-441">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-441">
         - HtmlCoercion</span></span><br><span data-ttu-id="672c7-442">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-442">
         - ImageCoercion</span></span><br><span data-ttu-id="672c7-443">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-443">
         - MatrixBindings</span></span><br><span data-ttu-id="672c7-444">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-444">
         - MatrixCoercion</span></span><br><span data-ttu-id="672c7-445">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-445">
         - OoxmlCoercion</span></span><br><span data-ttu-id="672c7-446">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="672c7-446">
         - PdfFile</span></span><br><span data-ttu-id="672c7-447">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-447">
         - Selection</span></span><br><span data-ttu-id="672c7-448">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-448">
         - Settings</span></span><br><span data-ttu-id="672c7-449">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-449">
         - TableBindings</span></span><br><span data-ttu-id="672c7-450">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-450">
         - TableCoercion</span></span><br><span data-ttu-id="672c7-451">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-451">
         - TextBindings</span></span><br><span data-ttu-id="672c7-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-452">
         - TextCoercion</span></span><br><span data-ttu-id="672c7-453">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="672c7-453">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-454">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="672c7-454">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="672c7-455">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-455">- TaskPane</span></span><br><span data-ttu-id="672c7-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="672c7-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="672c7-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="672c7-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="672c7-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="672c7-461">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-461">- BindingEvents</span></span><br><span data-ttu-id="672c7-462">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-462">
         - CompressedFile</span></span><br><span data-ttu-id="672c7-463">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="672c7-463">
         - CustomXmlParts</span></span><br><span data-ttu-id="672c7-464">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-464">
         - DocumentEvents</span></span><br><span data-ttu-id="672c7-465">
         - File</span><span class="sxs-lookup"><span data-stu-id="672c7-465">
         - File</span></span><br><span data-ttu-id="672c7-466">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-466">
         - HtmlCoercion</span></span><br><span data-ttu-id="672c7-467">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-467">
         - ImageCoercion</span></span><br><span data-ttu-id="672c7-468">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-468">
         - MatrixBindings</span></span><br><span data-ttu-id="672c7-469">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-469">
         - MatrixCoercion</span></span><br><span data-ttu-id="672c7-470">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-470">
         - OoxmlCoercion</span></span><br><span data-ttu-id="672c7-471">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="672c7-471">
         - PdfFile</span></span><br><span data-ttu-id="672c7-472">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-472">
         - Selection</span></span><br><span data-ttu-id="672c7-473">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-473">
         - Settings</span></span><br><span data-ttu-id="672c7-474">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-474">
         - TableBindings</span></span><br><span data-ttu-id="672c7-475">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-475">
         - TableCoercion</span></span><br><span data-ttu-id="672c7-476">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-476">
         - TextBindings</span></span><br><span data-ttu-id="672c7-477">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-477">
         - TextCoercion</span></span><br><span data-ttu-id="672c7-478">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="672c7-478">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-479">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="672c7-479">Office for iPad</span></span></td>
    <td> <span data-ttu-id="672c7-480">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-480">- TaskPane</span></span></td>
    <td> <span data-ttu-id="672c7-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="672c7-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="672c7-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="672c7-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="672c7-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="672c7-485">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-485">- BindingEvents</span></span><br><span data-ttu-id="672c7-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-486">
         - CompressedFile</span></span><br><span data-ttu-id="672c7-487">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="672c7-487">
         - CustomXmlParts</span></span><br><span data-ttu-id="672c7-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-488">
         - DocumentEvents</span></span><br><span data-ttu-id="672c7-489">
         - File</span><span class="sxs-lookup"><span data-stu-id="672c7-489">
         - File</span></span><br><span data-ttu-id="672c7-490">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-490">
         - HtmlCoercion</span></span><br><span data-ttu-id="672c7-491">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-491">
         - ImageCoercion</span></span><br><span data-ttu-id="672c7-492">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-492">
         - MatrixBindings</span></span><br><span data-ttu-id="672c7-493">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-493">
         - MatrixCoercion</span></span><br><span data-ttu-id="672c7-494">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-494">
         - OoxmlCoercion</span></span><br><span data-ttu-id="672c7-495">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="672c7-495">
         - PdfFile</span></span><br><span data-ttu-id="672c7-496">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-496">
         - Selection</span></span><br><span data-ttu-id="672c7-497">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-497">
         - Settings</span></span><br><span data-ttu-id="672c7-498">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-498">
         - TableBindings</span></span><br><span data-ttu-id="672c7-499">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-499">
         - TableCoercion</span></span><br><span data-ttu-id="672c7-500">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-500">
         - TextBindings</span></span><br><span data-ttu-id="672c7-501">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-501">
         - TextCoercion</span></span><br><span data-ttu-id="672c7-502">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="672c7-502">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-503">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="672c7-503">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="672c7-504">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-504">- TaskPane</span></span><br><span data-ttu-id="672c7-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="672c7-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="672c7-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="672c7-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="672c7-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="672c7-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="672c7-510">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-510">- BindingEvents</span></span><br><span data-ttu-id="672c7-511">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-511">
         - CompressedFile</span></span><br><span data-ttu-id="672c7-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="672c7-512">
         - CustomXmlParts</span></span><br><span data-ttu-id="672c7-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-513">
         - DocumentEvents</span></span><br><span data-ttu-id="672c7-514">
         - File</span><span class="sxs-lookup"><span data-stu-id="672c7-514">
         - File</span></span><br><span data-ttu-id="672c7-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-515">
         - HtmlCoercion</span></span><br><span data-ttu-id="672c7-516">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-516">
         - ImageCoercion</span></span><br><span data-ttu-id="672c7-517">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-517">
         - MatrixBindings</span></span><br><span data-ttu-id="672c7-518">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-518">
         - MatrixCoercion</span></span><br><span data-ttu-id="672c7-519">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-519">
         - OoxmlCoercion</span></span><br><span data-ttu-id="672c7-520">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="672c7-520">
         - PdfFile</span></span><br><span data-ttu-id="672c7-521">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-521">
         - Selection</span></span><br><span data-ttu-id="672c7-522">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-522">
         - Settings</span></span><br><span data-ttu-id="672c7-523">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-523">
         - TableBindings</span></span><br><span data-ttu-id="672c7-524">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-524">
         - TableCoercion</span></span><br><span data-ttu-id="672c7-525">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-525">
         - TextBindings</span></span><br><span data-ttu-id="672c7-526">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-526">
         - TextCoercion</span></span><br><span data-ttu-id="672c7-527">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="672c7-527">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-528">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="672c7-528">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="672c7-529">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-529">- TaskPane</span></span><br><span data-ttu-id="672c7-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="672c7-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="672c7-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="672c7-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="672c7-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="672c7-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="672c7-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="672c7-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="672c7-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-535">- BindingEvents</span></span><br><span data-ttu-id="672c7-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-536">
         - CompressedFile</span></span><br><span data-ttu-id="672c7-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="672c7-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="672c7-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-538">
         - DocumentEvents</span></span><br><span data-ttu-id="672c7-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="672c7-539">
         - File</span></span><br><span data-ttu-id="672c7-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="672c7-541">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-541">
         - ImageCoercion</span></span><br><span data-ttu-id="672c7-542">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-542">
         - MatrixBindings</span></span><br><span data-ttu-id="672c7-543">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-543">
         - MatrixCoercion</span></span><br><span data-ttu-id="672c7-544">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-544">
         - OoxmlCoercion</span></span><br><span data-ttu-id="672c7-545">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="672c7-545">
         - PdfFile</span></span><br><span data-ttu-id="672c7-546">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-546">
         - Selection</span></span><br><span data-ttu-id="672c7-547">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-547">
         - Settings</span></span><br><span data-ttu-id="672c7-548">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-548">
         - TableBindings</span></span><br><span data-ttu-id="672c7-549">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-549">
         - TableCoercion</span></span><br><span data-ttu-id="672c7-550">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="672c7-550">
         - TextBindings</span></span><br><span data-ttu-id="672c7-551">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-551">
         - TextCoercion</span></span><br><span data-ttu-id="672c7-552">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="672c7-552">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="672c7-553">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="672c7-553">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="672c7-554">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="672c7-554">Platform</span></span></th>
    <th><span data-ttu-id="672c7-555">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="672c7-555">Extension points</span></span></th>
    <th><span data-ttu-id="672c7-556">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="672c7-556">API requirement sets</span></span></th>
    <th><span data-ttu-id="672c7-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="672c7-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-558">Office Online</span><span class="sxs-lookup"><span data-stu-id="672c7-558">Office Online</span></span></td>
    <td> <span data-ttu-id="672c7-559">- Contenu</span><span class="sxs-lookup"><span data-stu-id="672c7-559">- Content</span></span><br><span data-ttu-id="672c7-560">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-560">
         - TaskPane</span></span><br><span data-ttu-id="672c7-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="672c7-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="672c7-563">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="672c7-563">- ActiveView</span></span><br><span data-ttu-id="672c7-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-564">
         - CompressedFile</span></span><br><span data-ttu-id="672c7-565">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-565">
         - DocumentEvents</span></span><br><span data-ttu-id="672c7-566">
         - File</span><span class="sxs-lookup"><span data-stu-id="672c7-566">
         - File</span></span><br><span data-ttu-id="672c7-567">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-567">
         - ImageCoercion</span></span><br><span data-ttu-id="672c7-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="672c7-568">
         - PdfFile</span></span><br><span data-ttu-id="672c7-569">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-569">
         - Selection</span></span><br><span data-ttu-id="672c7-570">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-570">
         - Settings</span></span><br><span data-ttu-id="672c7-571">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-571">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-572">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="672c7-572">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="672c7-573">- Contenu</span><span class="sxs-lookup"><span data-stu-id="672c7-573">- Content</span></span><br><span data-ttu-id="672c7-574">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-574">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="672c7-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="672c7-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="672c7-576">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="672c7-576">- ActiveView</span></span><br><span data-ttu-id="672c7-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-577">
         - CompressedFile</span></span><br><span data-ttu-id="672c7-578">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-578">
         - DocumentEvents</span></span><br><span data-ttu-id="672c7-579">
         - File</span><span class="sxs-lookup"><span data-stu-id="672c7-579">
         - File</span></span><br><span data-ttu-id="672c7-580">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-580">
         - ImageCoercion</span></span><br><span data-ttu-id="672c7-581">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="672c7-581">
         - PdfFile</span></span><br><span data-ttu-id="672c7-582">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-582">
         - Selection</span></span><br><span data-ttu-id="672c7-583">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-583">
         - Settings</span></span><br><span data-ttu-id="672c7-584">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-584">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-585">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="672c7-585">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="672c7-586">- Contenu</span><span class="sxs-lookup"><span data-stu-id="672c7-586">- Content</span></span><br><span data-ttu-id="672c7-587">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-587">
         - TaskPane</span></span><br><span data-ttu-id="672c7-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="672c7-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="672c7-590">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="672c7-590">- ActiveView</span></span><br><span data-ttu-id="672c7-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-591">
         - CompressedFile</span></span><br><span data-ttu-id="672c7-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-592">
         - DocumentEvents</span></span><br><span data-ttu-id="672c7-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="672c7-593">
         - File</span></span><br><span data-ttu-id="672c7-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-594">
         - ImageCoercion</span></span><br><span data-ttu-id="672c7-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="672c7-595">
         - PdfFile</span></span><br><span data-ttu-id="672c7-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-596">
         - Selection</span></span><br><span data-ttu-id="672c7-597">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-597">
         - Settings</span></span><br><span data-ttu-id="672c7-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-598">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-599">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="672c7-599">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="672c7-600">- Contenu</span><span class="sxs-lookup"><span data-stu-id="672c7-600">- Content</span></span><br><span data-ttu-id="672c7-601">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-601">
         - TaskPane</span></span><br><span data-ttu-id="672c7-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="672c7-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="672c7-604">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="672c7-604">- ActiveView</span></span><br><span data-ttu-id="672c7-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-605">
         - CompressedFile</span></span><br><span data-ttu-id="672c7-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-606">
         - DocumentEvents</span></span><br><span data-ttu-id="672c7-607">
         - File</span><span class="sxs-lookup"><span data-stu-id="672c7-607">
         - File</span></span><br><span data-ttu-id="672c7-608">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-608">
         - ImageCoercion</span></span><br><span data-ttu-id="672c7-609">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="672c7-609">
         - PdfFile</span></span><br><span data-ttu-id="672c7-610">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-610">
         - Selection</span></span><br><span data-ttu-id="672c7-611">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-611">
         - Settings</span></span><br><span data-ttu-id="672c7-612">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-612">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-613">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="672c7-613">Office for iPad</span></span></td>
    <td> <span data-ttu-id="672c7-614">- Contenu</span><span class="sxs-lookup"><span data-stu-id="672c7-614">- Content</span></span><br><span data-ttu-id="672c7-615">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-615">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="672c7-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="672c7-617">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="672c7-617">- ActiveView</span></span><br><span data-ttu-id="672c7-618">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-618">
         - CompressedFile</span></span><br><span data-ttu-id="672c7-619">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-619">
         - DocumentEvents</span></span><br><span data-ttu-id="672c7-620">
         - File</span><span class="sxs-lookup"><span data-stu-id="672c7-620">
         - File</span></span><br><span data-ttu-id="672c7-621">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="672c7-621">
         - PdfFile</span></span><br><span data-ttu-id="672c7-622">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-622">
         - Selection</span></span><br><span data-ttu-id="672c7-623">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-623">
         - Settings</span></span><br><span data-ttu-id="672c7-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-624">
         - TextCoercion</span></span><br><span data-ttu-id="672c7-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-625">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-626">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="672c7-626">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="672c7-627">- Contenu</span><span class="sxs-lookup"><span data-stu-id="672c7-627">- Content</span></span><br><span data-ttu-id="672c7-628">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-628">
         - TaskPane</span></span><br><span data-ttu-id="672c7-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="672c7-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="672c7-631">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="672c7-631">- ActiveView</span></span><br><span data-ttu-id="672c7-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-632">
         - CompressedFile</span></span><br><span data-ttu-id="672c7-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-633">
         - DocumentEvents</span></span><br><span data-ttu-id="672c7-634">
         - File</span><span class="sxs-lookup"><span data-stu-id="672c7-634">
         - File</span></span><br><span data-ttu-id="672c7-635">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-635">
         - ImageCoercion</span></span><br><span data-ttu-id="672c7-636">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="672c7-636">
         - PdfFile</span></span><br><span data-ttu-id="672c7-637">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-637">
         - Selection</span></span><br><span data-ttu-id="672c7-638">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-638">
         - Settings</span></span><br><span data-ttu-id="672c7-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-639">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-640">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="672c7-640">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="672c7-641">- Contenu</span><span class="sxs-lookup"><span data-stu-id="672c7-641">- Content</span></span><br><span data-ttu-id="672c7-642">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-642">
         - TaskPane</span></span><br><span data-ttu-id="672c7-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="672c7-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="672c7-645">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="672c7-645">- ActiveView</span></span><br><span data-ttu-id="672c7-646">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="672c7-646">
         - CompressedFile</span></span><br><span data-ttu-id="672c7-647">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-647">
         - DocumentEvents</span></span><br><span data-ttu-id="672c7-648">
         - File</span><span class="sxs-lookup"><span data-stu-id="672c7-648">
         - File</span></span><br><span data-ttu-id="672c7-649">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-649">
         - ImageCoercion</span></span><br><span data-ttu-id="672c7-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="672c7-650">
         - PdfFile</span></span><br><span data-ttu-id="672c7-651">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-651">
         - Selection</span></span><br><span data-ttu-id="672c7-652">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-652">
         - Settings</span></span><br><span data-ttu-id="672c7-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-653">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="672c7-654">OneNote</span><span class="sxs-lookup"><span data-stu-id="672c7-654">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="672c7-655">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="672c7-655">Platform</span></span></th>
    <th><span data-ttu-id="672c7-656">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="672c7-656">Extension points</span></span></th>
    <th><span data-ttu-id="672c7-657">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="672c7-657">API requirement sets</span></span></th>
    <th><span data-ttu-id="672c7-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="672c7-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-659">Office Online</span><span class="sxs-lookup"><span data-stu-id="672c7-659">Office Online</span></span></td>
    <td> <span data-ttu-id="672c7-660">- Contenu</span><span class="sxs-lookup"><span data-stu-id="672c7-660">- Content</span></span><br><span data-ttu-id="672c7-661">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-661">
         - TaskPane</span></span><br><span data-ttu-id="672c7-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="672c7-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="672c7-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="672c7-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="672c7-665">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="672c7-665">- DocumentEvents</span></span><br><span data-ttu-id="672c7-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="672c7-667">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-667">
         - ImageCoercion</span></span><br><span data-ttu-id="672c7-668">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="672c7-668">
         - Settings</span></span><br><span data-ttu-id="672c7-669">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-669">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="672c7-670">Projet</span><span class="sxs-lookup"><span data-stu-id="672c7-670">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="672c7-671">Plateforme</span><span class="sxs-lookup"><span data-stu-id="672c7-671">Platform</span></span></th>
    <th><span data-ttu-id="672c7-672">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="672c7-672">Extension points</span></span></th>
    <th><span data-ttu-id="672c7-673">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="672c7-673">API requirement sets</span></span></th>
    <th><span data-ttu-id="672c7-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="672c7-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-675">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="672c7-675">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="672c7-676">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-676">- TaskPane</span></span></td>
    <td> <span data-ttu-id="672c7-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="672c7-678">- Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-678">- Selection</span></span><br><span data-ttu-id="672c7-679">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-679">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-680">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="672c7-680">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="672c7-681">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-681">- TaskPane</span></span></td>
    <td> <span data-ttu-id="672c7-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="672c7-683">- Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-683">- Selection</span></span><br><span data-ttu-id="672c7-684">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-684">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="672c7-685">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="672c7-685">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="672c7-686">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="672c7-686">- TaskPane</span></span></td>
    <td> <span data-ttu-id="672c7-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="672c7-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="672c7-688">- Selection</span><span class="sxs-lookup"><span data-stu-id="672c7-688">- Selection</span></span><br><span data-ttu-id="672c7-689">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="672c7-689">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="672c7-690">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="672c7-690">See also</span></span>

- [<span data-ttu-id="672c7-691">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="672c7-691">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="672c7-692">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="672c7-692">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="672c7-693">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="672c7-693">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="672c7-694">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="672c7-694">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
