---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint, OneNote et Project.
ms.date: 11/07/2018
ms.openlocfilehash: f8d7d9d393531301829b31dd171a5332a0da536b
ms.sourcegitcommit: 9b021af6cb23a58486d6c5c7492be425e309bea1
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/15/2018
ms.locfileid: "26533797"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="033b0-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="033b0-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="033b0-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span><span class="sxs-lookup"><span data-stu-id="033b0-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="033b0-p102">Le numéro de build pour Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que les ensembles de conditions requises des API communes, ExcelApi 1.1 et WordApi 1.1.</span><span class="sxs-lookup"><span data-stu-id="033b0-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="033b0-108">Excel</span><span class="sxs-lookup"><span data-stu-id="033b0-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="033b0-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="033b0-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="033b0-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="033b0-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="033b0-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="033b0-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="033b0-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="033b0-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="033b0-113">Office Online</span></span></td>
    <td> <span data-ttu-id="033b0-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-114">- Taskpane</span></span><br><span data-ttu-id="033b0-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="033b0-115">
        - Content</span></span><br><span data-ttu-id="033b0-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="033b0-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="033b0-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="033b0-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="033b0-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="033b0-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="033b0-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="033b0-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="033b0-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="033b0-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="033b0-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="033b0-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="033b0-123">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="033b0-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="033b0-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="033b0-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="033b0-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-126">
        -BindingEvents</span></span><br><span data-ttu-id="033b0-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-127">
        -CompressedFile</span></span><br><span data-ttu-id="033b0-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-128">
        -DocumentEvents</span></span><br><span data-ttu-id="033b0-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="033b0-129">
        - File</span></span><br><span data-ttu-id="033b0-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-130">
        -MatrixBindings</span></span><br><span data-ttu-id="033b0-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-131">
        -MatrixCoercion</span></span><br><span data-ttu-id="033b0-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-132">
        - Selection</span></span><br><span data-ttu-id="033b0-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-133">
        - Settings</span></span><br><span data-ttu-id="033b0-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-134">
        -TableBindings</span></span><br><span data-ttu-id="033b0-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-135">
        -TableCoercion</span></span><br><span data-ttu-id="033b0-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-136">
        -TextBindings</span></span><br><span data-ttu-id="033b0-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-137">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-138">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="033b0-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="033b0-139">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-139">
        - Taskpane</span></span><br><span data-ttu-id="033b0-140">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="033b0-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="033b0-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="033b0-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-142">
        -BindingEvents</span></span><br><span data-ttu-id="033b0-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-143">
        -CompressedFile</span></span><br><span data-ttu-id="033b0-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-144">
        -DocumentEvents</span></span><br><span data-ttu-id="033b0-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="033b0-145">
        - File</span></span><br><span data-ttu-id="033b0-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-146">
        -ImageCoercion</span></span><br><span data-ttu-id="033b0-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-147">
        -MatrixBindings</span></span><br><span data-ttu-id="033b0-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-148">
        -MatrixCoercion</span></span><br><span data-ttu-id="033b0-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-149">
        - Selection</span></span><br><span data-ttu-id="033b0-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-150">
        - Settings</span></span><br><span data-ttu-id="033b0-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-151">
        -TableBindings</span></span><br><span data-ttu-id="033b0-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-152">
        -TableCoercion</span></span><br><span data-ttu-id="033b0-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-153">
        -TextBindings</span></span><br><span data-ttu-id="033b0-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-154">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-155">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="033b0-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="033b0-156">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-156">- Taskpane</span></span><br><span data-ttu-id="033b0-157">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="033b0-157">
        - Content</span></span><br><span data-ttu-id="033b0-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="033b0-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="033b0-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="033b0-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="033b0-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="033b0-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="033b0-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="033b0-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="033b0-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="033b0-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="033b0-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="033b0-165">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="033b0-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="033b0-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="033b0-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="033b0-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-168">-BindingEvents</span></span><br><span data-ttu-id="033b0-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-169">
        -CompressedFile</span></span><br><span data-ttu-id="033b0-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-170">
        -DocumentEvents</span></span><br><span data-ttu-id="033b0-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="033b0-171">
        - File</span></span><br><span data-ttu-id="033b0-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-172">
        -ImageCoercion</span></span><br><span data-ttu-id="033b0-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-173">
        -MatrixBindings</span></span><br><span data-ttu-id="033b0-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="033b0-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-175">
        - Selection</span></span><br><span data-ttu-id="033b0-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-176">
        - Settings</span></span><br><span data-ttu-id="033b0-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-177">
        -TableBindings</span></span><br><span data-ttu-id="033b0-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-178">
        -TableCoercion</span></span><br><span data-ttu-id="033b0-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-179">
        -TextBindings</span></span><br><span data-ttu-id="033b0-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-181">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="033b0-181">Office for Windows Desktop.</span></span></td>
    <td><span data-ttu-id="033b0-182">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-182">- Taskpane</span></span><br><span data-ttu-id="033b0-183">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="033b0-183">
        - Content</span></span><br><span data-ttu-id="033b0-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="033b0-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="033b0-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="033b0-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="033b0-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="033b0-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="033b0-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="033b0-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="033b0-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="033b0-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="033b0-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="033b0-191">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="033b0-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="033b0-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="033b0-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="033b0-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-194">-BindingEvents</span></span><br><span data-ttu-id="033b0-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-195">
        -CompressedFile</span></span><br><span data-ttu-id="033b0-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-196">
        -DocumentEvents</span></span><br><span data-ttu-id="033b0-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="033b0-197">
        - File</span></span><br><span data-ttu-id="033b0-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-198">
        -ImageCoercion</span></span><br><span data-ttu-id="033b0-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-199">
        -MatrixBindings</span></span><br><span data-ttu-id="033b0-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-200">
        -MatrixCoercion</span></span><br><span data-ttu-id="033b0-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-201">
        - Selection</span></span><br><span data-ttu-id="033b0-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-202">
        - Settings</span></span><br><span data-ttu-id="033b0-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-203">
        -TableBindings</span></span><br><span data-ttu-id="033b0-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-204">
        -TableCoercion</span></span><br><span data-ttu-id="033b0-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-205">
        -TextBindings</span></span><br><span data-ttu-id="033b0-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-206">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-207">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="033b0-207">Office for iOS</span></span></td>
    <td><span data-ttu-id="033b0-208">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-208">- Taskpane</span></span><br><span data-ttu-id="033b0-209">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="033b0-209">
        - Content</span></span></td>
    <td><span data-ttu-id="033b0-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="033b0-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="033b0-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="033b0-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="033b0-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="033b0-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="033b0-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="033b0-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="033b0-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="033b0-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="033b0-216">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="033b0-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="033b0-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="033b0-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="033b0-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-219">-BindingEvents</span></span><br><span data-ttu-id="033b0-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-220">
        -CompressedFile</span></span><br><span data-ttu-id="033b0-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-221">
        -DocumentEvents</span></span><br><span data-ttu-id="033b0-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="033b0-222">
        - File</span></span><br><span data-ttu-id="033b0-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-223">
        -ImageCoercion</span></span><br><span data-ttu-id="033b0-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-224">
        -MatrixBindings</span></span><br><span data-ttu-id="033b0-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-225">
        -MatrixCoercion</span></span><br><span data-ttu-id="033b0-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-226">
        - Selection</span></span><br><span data-ttu-id="033b0-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-227">
        - Settings</span></span><br><span data-ttu-id="033b0-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-228">
        -TableBindings</span></span><br><span data-ttu-id="033b0-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-229">
        -TableCoercion</span></span><br><span data-ttu-id="033b0-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-230">
        -TextBindings</span></span><br><span data-ttu-id="033b0-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-231">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-232">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="033b0-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="033b0-233">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-233">- Taskpane</span></span><br><span data-ttu-id="033b0-234">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="033b0-234">
        - Content</span></span><br><span data-ttu-id="033b0-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="033b0-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="033b0-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="033b0-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="033b0-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="033b0-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="033b0-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="033b0-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="033b0-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="033b0-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="033b0-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="033b0-242">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="033b0-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="033b0-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="033b0-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="033b0-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-245">-BindingEvents</span></span><br><span data-ttu-id="033b0-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-246">
        -CompressedFile</span></span><br><span data-ttu-id="033b0-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-247">
        -DocumentEvents</span></span><br><span data-ttu-id="033b0-248">
        - File</span><span class="sxs-lookup"><span data-stu-id="033b0-248">
        - File</span></span><br><span data-ttu-id="033b0-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-249">
        -ImageCoercion</span></span><br><span data-ttu-id="033b0-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-250">
        -MatrixBindings</span></span><br><span data-ttu-id="033b0-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-251">
        -MatrixCoercion</span></span><br><span data-ttu-id="033b0-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="033b0-252">
        -PdfFile</span></span><br><span data-ttu-id="033b0-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-253">
        - Selection</span></span><br><span data-ttu-id="033b0-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-254">
        - Settings</span></span><br><span data-ttu-id="033b0-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-255">
        -TableBindings</span></span><br><span data-ttu-id="033b0-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-256">
        -TableCoercion</span></span><br><span data-ttu-id="033b0-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-257">
        -TextBindings</span></span><br><span data-ttu-id="033b0-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-258">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-259">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="033b0-259">Office for Mac</span></span></td>
    <td><span data-ttu-id="033b0-260">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-260">- Taskpane</span></span><br><span data-ttu-id="033b0-261">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="033b0-261">
        - Content</span></span><br><span data-ttu-id="033b0-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="033b0-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="033b0-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="033b0-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="033b0-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="033b0-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="033b0-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="033b0-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="033b0-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="033b0-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="033b0-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="033b0-269">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="033b0-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="033b0-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="033b0-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="033b0-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-272">-BindingEvents</span></span><br><span data-ttu-id="033b0-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-273">
        -CompressedFile</span></span><br><span data-ttu-id="033b0-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-274">
        -DocumentEvents</span></span><br><span data-ttu-id="033b0-275">
        - File</span><span class="sxs-lookup"><span data-stu-id="033b0-275">
        - File</span></span><br><span data-ttu-id="033b0-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-276">
        -ImageCoercion</span></span><br><span data-ttu-id="033b0-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-277">
        -MatrixBindings</span></span><br><span data-ttu-id="033b0-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-278">
        -MatrixCoercion</span></span><br><span data-ttu-id="033b0-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="033b0-279">
        -PdfFile</span></span><br><span data-ttu-id="033b0-280">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-280">
        - Selection</span></span><br><span data-ttu-id="033b0-281">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-281">
        - Settings</span></span><br><span data-ttu-id="033b0-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-282">
        -TableBindings</span></span><br><span data-ttu-id="033b0-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-283">
        -TableCoercion</span></span><br><span data-ttu-id="033b0-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-284">
        -TextBindings</span></span><br><span data-ttu-id="033b0-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-285">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="033b0-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="033b0-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="033b0-287">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="033b0-287">Platform</span></span></th>
    <th><span data-ttu-id="033b0-288">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="033b0-288">Extension points</span></span></th>
    <th><span data-ttu-id="033b0-289">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="033b0-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="033b0-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="033b0-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="033b0-291">Office Online</span></span></td>
    <td> <span data-ttu-id="033b0-292">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="033b0-292">- Mail Read</span></span><br><span data-ttu-id="033b0-293">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="033b0-293">
      - Mail Compose</span></span><br><span data-ttu-id="033b0-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="033b0-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="033b0-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="033b0-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="033b0-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="033b0-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="033b0-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="033b0-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="033b0-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="033b0-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="033b0-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="033b0-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="033b0-302">Non disponible</span><span class="sxs-lookup"><span data-stu-id="033b0-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-303">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="033b0-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="033b0-304">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="033b0-304">- Mail Read</span></span><br><span data-ttu-id="033b0-305">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="033b0-305">
      - Mail Compose</span></span><br><span data-ttu-id="033b0-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="033b0-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="033b0-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="033b0-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="033b0-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="033b0-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="033b0-311">Non disponible</span><span class="sxs-lookup"><span data-stu-id="033b0-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-312">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="033b0-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="033b0-313">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="033b0-313">- Mail Read</span></span><br><span data-ttu-id="033b0-314">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="033b0-314">
      - Mail Compose</span></span><br><span data-ttu-id="033b0-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="033b0-316">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="033b0-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="033b0-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="033b0-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="033b0-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="033b0-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="033b0-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="033b0-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="033b0-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="033b0-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="033b0-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="033b0-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="033b0-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="033b0-324">Non disponible</span><span class="sxs-lookup"><span data-stu-id="033b0-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-325">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="033b0-325">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="033b0-326">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="033b0-326">- Mail Read</span></span><br><span data-ttu-id="033b0-327">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="033b0-327">
      - Mail Compose</span></span><br><span data-ttu-id="033b0-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="033b0-329">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="033b0-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="033b0-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="033b0-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="033b0-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="033b0-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="033b0-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="033b0-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="033b0-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="033b0-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="033b0-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="033b0-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="033b0-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="033b0-337">Non disponible</span><span class="sxs-lookup"><span data-stu-id="033b0-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-338">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="033b0-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="033b0-339">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="033b0-339">- Mail Read</span></span><br><span data-ttu-id="033b0-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="033b0-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="033b0-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="033b0-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="033b0-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="033b0-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="033b0-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="033b0-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="033b0-346">Non disponible</span><span class="sxs-lookup"><span data-stu-id="033b0-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-347">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="033b0-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="033b0-348">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="033b0-348">- Mail Read</span></span><br><span data-ttu-id="033b0-349">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="033b0-349">
      - Mail Compose</span></span><br><span data-ttu-id="033b0-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="033b0-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="033b0-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="033b0-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="033b0-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="033b0-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="033b0-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="033b0-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="033b0-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="033b0-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="033b0-357">Non disponible</span><span class="sxs-lookup"><span data-stu-id="033b0-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-358">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="033b0-358">Office for Mac</span></span></td>
    <td> <span data-ttu-id="033b0-359">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="033b0-359">- Mail Read</span></span><br><span data-ttu-id="033b0-360">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="033b0-360">
      - Mail Compose</span></span><br><span data-ttu-id="033b0-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="033b0-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="033b0-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="033b0-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="033b0-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="033b0-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="033b0-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="033b0-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="033b0-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="033b0-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="033b0-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="033b0-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="033b0-369">Non disponible</span><span class="sxs-lookup"><span data-stu-id="033b0-369">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-370">Office pour Android</span><span class="sxs-lookup"><span data-stu-id="033b0-370">Office for Android</span></span></td>
    <td> <span data-ttu-id="033b0-371">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="033b0-371">- Mail Read</span></span><br><span data-ttu-id="033b0-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="033b0-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="033b0-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="033b0-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="033b0-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="033b0-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="033b0-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="033b0-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="033b0-378">Non disponible</span><span class="sxs-lookup"><span data-stu-id="033b0-378">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="033b0-379">Word</span><span class="sxs-lookup"><span data-stu-id="033b0-379">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="033b0-380">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="033b0-380">Platform</span></span></th>
    <th><span data-ttu-id="033b0-381">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="033b0-381">Extension points</span></span></th>
    <th><span data-ttu-id="033b0-382">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="033b0-382">API requirement sets</span></span></th>
    <th><span data-ttu-id="033b0-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="033b0-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-384">Office Online</span><span class="sxs-lookup"><span data-stu-id="033b0-384">Office Online</span></span></td>
    <td> <span data-ttu-id="033b0-385">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-385">- Taskpane</span></span><br><span data-ttu-id="033b0-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="033b0-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="033b0-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="033b0-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="033b0-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="033b0-391">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-391">-BindingEvents</span></span><br><span data-ttu-id="033b0-392">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="033b0-392">
         -</span></span><br><span data-ttu-id="033b0-393">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-393">
         -DocumentEvents</span></span><br><span data-ttu-id="033b0-394">
         - File</span><span class="sxs-lookup"><span data-stu-id="033b0-394">
         - File</span></span><br><span data-ttu-id="033b0-395">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-395">
         -HtmlCoercion</span></span><br><span data-ttu-id="033b0-396">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-396">
         -ImageCoercion</span></span><br><span data-ttu-id="033b0-397">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-397">
         -MatrixBindings</span></span><br><span data-ttu-id="033b0-398">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-398">
         -MatrixCoercion</span></span><br><span data-ttu-id="033b0-399">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-399">
         -OoxmlCoercion</span></span><br><span data-ttu-id="033b0-400">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="033b0-400">
         -PdfFile</span></span><br><span data-ttu-id="033b0-401">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-401">
         - Selection</span></span><br><span data-ttu-id="033b0-402">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-402">
         - Settings</span></span><br><span data-ttu-id="033b0-403">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-403">
         -TableBindings</span></span><br><span data-ttu-id="033b0-404">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-404">
         -TableCoercion</span></span><br><span data-ttu-id="033b0-405">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-405">
         -TextBindings</span></span><br><span data-ttu-id="033b0-406">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-406">
         -TextCoercion</span></span><br><span data-ttu-id="033b0-407">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="033b0-407">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-408">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="033b0-408">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="033b0-409">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-409">- Taskpane</span></span></td>
    <td> <span data-ttu-id="033b0-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="033b0-411">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-411">-BindingEvents</span></span><br><span data-ttu-id="033b0-412">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-412">
         -CompressedFile</span></span><br><span data-ttu-id="033b0-413">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="033b0-413">
         -</span></span><br><span data-ttu-id="033b0-414">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-414">
         -DocumentEvents</span></span><br><span data-ttu-id="033b0-415">
         - File</span><span class="sxs-lookup"><span data-stu-id="033b0-415">
         - File</span></span><br><span data-ttu-id="033b0-416">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-416">
         -HtmlCoercion</span></span><br><span data-ttu-id="033b0-417">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-417">
         -ImageCoercion</span></span><br><span data-ttu-id="033b0-418">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-418">
         -MatrixBindings</span></span><br><span data-ttu-id="033b0-419">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-419">
         -MatrixCoercion</span></span><br><span data-ttu-id="033b0-420">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-420">
         -OoxmlCoercion</span></span><br><span data-ttu-id="033b0-421">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="033b0-421">
         -PdfFile</span></span><br><span data-ttu-id="033b0-422">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-422">
         - Selection</span></span><br><span data-ttu-id="033b0-423">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-423">
         - Settings</span></span><br><span data-ttu-id="033b0-424">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-424">
         -TableBindings</span></span><br><span data-ttu-id="033b0-425">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-425">
         -TableCoercion</span></span><br><span data-ttu-id="033b0-426">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-426">
         -TextBindings</span></span><br><span data-ttu-id="033b0-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-427">
         -TextCoercion</span></span><br><span data-ttu-id="033b0-428">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="033b0-428">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-429">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="033b0-429">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="033b0-430">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-430">- Taskpane</span></span><br><span data-ttu-id="033b0-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="033b0-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="033b0-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="033b0-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="033b0-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="033b0-436">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-436">-BindingEvents</span></span><br><span data-ttu-id="033b0-437">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-437">
         -CompressedFile</span></span><br><span data-ttu-id="033b0-438">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="033b0-438">
         -</span></span><br><span data-ttu-id="033b0-439">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-439">
         -DocumentEvents</span></span><br><span data-ttu-id="033b0-440">
         - File</span><span class="sxs-lookup"><span data-stu-id="033b0-440">
         - File</span></span><br><span data-ttu-id="033b0-441">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-441">
         -HtmlCoercion</span></span><br><span data-ttu-id="033b0-442">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-442">
         -ImageCoercion</span></span><br><span data-ttu-id="033b0-443">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-443">
         -MatrixBindings</span></span><br><span data-ttu-id="033b0-444">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-444">
         -MatrixCoercion</span></span><br><span data-ttu-id="033b0-445">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-445">
         -OoxmlCoercion</span></span><br><span data-ttu-id="033b0-446">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="033b0-446">
         -PdfFile</span></span><br><span data-ttu-id="033b0-447">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-447">
         - Selection</span></span><br><span data-ttu-id="033b0-448">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-448">
         - Settings</span></span><br><span data-ttu-id="033b0-449">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-449">
         -TableBindings</span></span><br><span data-ttu-id="033b0-450">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-450">
         -TableCoercion</span></span><br><span data-ttu-id="033b0-451">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-451">
         -TextBindings</span></span><br><span data-ttu-id="033b0-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-452">
         -TextCoercion</span></span><br><span data-ttu-id="033b0-453">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="033b0-453">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-454">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="033b0-454">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="033b0-455">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-455">- Taskpane</span></span><br><span data-ttu-id="033b0-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="033b0-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="033b0-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="033b0-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="033b0-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="033b0-461">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-461">-BindingEvents</span></span><br><span data-ttu-id="033b0-462">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-462">
         -CompressedFile</span></span><br><span data-ttu-id="033b0-463">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="033b0-463">
         -</span></span><br><span data-ttu-id="033b0-464">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-464">
         -DocumentEvents</span></span><br><span data-ttu-id="033b0-465">
         - File</span><span class="sxs-lookup"><span data-stu-id="033b0-465">
         - File</span></span><br><span data-ttu-id="033b0-466">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-466">
         -HtmlCoercion</span></span><br><span data-ttu-id="033b0-467">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-467">
         -ImageCoercion</span></span><br><span data-ttu-id="033b0-468">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-468">
         -MatrixBindings</span></span><br><span data-ttu-id="033b0-469">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-469">
         -MatrixCoercion</span></span><br><span data-ttu-id="033b0-470">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-470">
         -OoxmlCoercion</span></span><br><span data-ttu-id="033b0-471">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="033b0-471">
         -PdfFile</span></span><br><span data-ttu-id="033b0-472">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-472">
         - Selection</span></span><br><span data-ttu-id="033b0-473">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-473">
         - Settings</span></span><br><span data-ttu-id="033b0-474">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-474">
         -TableBindings</span></span><br><span data-ttu-id="033b0-475">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-475">
         -TableCoercion</span></span><br><span data-ttu-id="033b0-476">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-476">
         -TextBindings</span></span><br><span data-ttu-id="033b0-477">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-477">
         -TextCoercion</span></span><br><span data-ttu-id="033b0-478">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="033b0-478">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-479">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="033b0-479">Office for iOS</span></span></td>
    <td> <span data-ttu-id="033b0-480">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-480">- Taskpane</span></span></td>
    <td> <span data-ttu-id="033b0-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="033b0-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="033b0-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="033b0-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="033b0-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="033b0-485">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-485">-BindingEvents</span></span><br><span data-ttu-id="033b0-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-486">
         -CompressedFile</span></span><br><span data-ttu-id="033b0-487">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="033b0-487">
         -</span></span><br><span data-ttu-id="033b0-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-488">
         -DocumentEvents</span></span><br><span data-ttu-id="033b0-489">
         - File</span><span class="sxs-lookup"><span data-stu-id="033b0-489">
         - File</span></span><br><span data-ttu-id="033b0-490">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-490">
         -HtmlCoercion</span></span><br><span data-ttu-id="033b0-491">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-491">
         -ImageCoercion</span></span><br><span data-ttu-id="033b0-492">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-492">
         -MatrixBindings</span></span><br><span data-ttu-id="033b0-493">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-493">
         -MatrixCoercion</span></span><br><span data-ttu-id="033b0-494">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-494">
         -OoxmlCoercion</span></span><br><span data-ttu-id="033b0-495">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="033b0-495">
         -PdfFile</span></span><br><span data-ttu-id="033b0-496">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-496">
         - Selection</span></span><br><span data-ttu-id="033b0-497">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-497">
         - Settings</span></span><br><span data-ttu-id="033b0-498">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-498">
         -TableBindings</span></span><br><span data-ttu-id="033b0-499">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-499">
         -TableCoercion</span></span><br><span data-ttu-id="033b0-500">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-500">
         -TextBindings</span></span><br><span data-ttu-id="033b0-501">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-501">
         -TextCoercion</span></span><br><span data-ttu-id="033b0-502">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="033b0-502">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-503">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="033b0-503">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="033b0-504">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-504">- Taskpane</span></span><br><span data-ttu-id="033b0-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="033b0-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="033b0-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="033b0-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="033b0-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="033b0-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="033b0-510">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-510">-BindingEvents</span></span><br><span data-ttu-id="033b0-511">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-511">
         -CompressedFile</span></span><br><span data-ttu-id="033b0-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="033b0-512">
         -</span></span><br><span data-ttu-id="033b0-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-513">
         -DocumentEvents</span></span><br><span data-ttu-id="033b0-514">
         - File</span><span class="sxs-lookup"><span data-stu-id="033b0-514">
         - File</span></span><br><span data-ttu-id="033b0-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-515">
         -HtmlCoercion</span></span><br><span data-ttu-id="033b0-516">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-516">
         -ImageCoercion</span></span><br><span data-ttu-id="033b0-517">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-517">
         -MatrixBindings</span></span><br><span data-ttu-id="033b0-518">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-518">
         -MatrixCoercion</span></span><br><span data-ttu-id="033b0-519">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-519">
         -OoxmlCoercion</span></span><br><span data-ttu-id="033b0-520">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="033b0-520">
         -PdfFile</span></span><br><span data-ttu-id="033b0-521">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-521">
         - Selection</span></span><br><span data-ttu-id="033b0-522">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-522">
         - Settings</span></span><br><span data-ttu-id="033b0-523">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-523">
         -TableBindings</span></span><br><span data-ttu-id="033b0-524">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-524">
         -TableCoercion</span></span><br><span data-ttu-id="033b0-525">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-525">
         -TextBindings</span></span><br><span data-ttu-id="033b0-526">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-526">
         -TextCoercion</span></span><br><span data-ttu-id="033b0-527">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="033b0-527">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-528">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="033b0-528">Office for Mac</span></span></td>
    <td> <span data-ttu-id="033b0-529">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-529">- Taskpane</span></span><br><span data-ttu-id="033b0-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="033b0-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="033b0-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="033b0-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="033b0-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="033b0-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="033b0-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="033b0-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="033b0-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-535">-BindingEvents</span></span><br><span data-ttu-id="033b0-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-536">
         -CompressedFile</span></span><br><span data-ttu-id="033b0-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="033b0-537">
         -</span></span><br><span data-ttu-id="033b0-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-538">
         -DocumentEvents</span></span><br><span data-ttu-id="033b0-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="033b0-539">
         - File</span></span><br><span data-ttu-id="033b0-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-540">
         -HtmlCoercion</span></span><br><span data-ttu-id="033b0-541">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-541">
         -ImageCoercion</span></span><br><span data-ttu-id="033b0-542">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-542">
         -MatrixBindings</span></span><br><span data-ttu-id="033b0-543">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-543">
         -MatrixCoercion</span></span><br><span data-ttu-id="033b0-544">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-544">
         -OoxmlCoercion</span></span><br><span data-ttu-id="033b0-545">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="033b0-545">
         -PdfFile</span></span><br><span data-ttu-id="033b0-546">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-546">
         - Selection</span></span><br><span data-ttu-id="033b0-547">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-547">
         - Settings</span></span><br><span data-ttu-id="033b0-548">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-548">
         -TableBindings</span></span><br><span data-ttu-id="033b0-549">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-549">
         -TableCoercion</span></span><br><span data-ttu-id="033b0-550">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="033b0-550">
         -TextBindings</span></span><br><span data-ttu-id="033b0-551">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-551">
         -TextCoercion</span></span><br><span data-ttu-id="033b0-552">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="033b0-552">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="033b0-553">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="033b0-553">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="033b0-554">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="033b0-554">Platform</span></span></th>
    <th><span data-ttu-id="033b0-555">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="033b0-555">Extension points</span></span></th>
    <th><span data-ttu-id="033b0-556">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="033b0-556">API requirement sets</span></span></th>
    <th><span data-ttu-id="033b0-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="033b0-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-558">Office Online</span><span class="sxs-lookup"><span data-stu-id="033b0-558">Office Online</span></span></td>
    <td> <span data-ttu-id="033b0-559">- Contenu</span><span class="sxs-lookup"><span data-stu-id="033b0-559">- Content</span></span><br><span data-ttu-id="033b0-560">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-560">
         - Taskpane</span></span><br><span data-ttu-id="033b0-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="033b0-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="033b0-563">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="033b0-563">-ActiveView</span></span><br><span data-ttu-id="033b0-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-564">
         -CompressedFile</span></span><br><span data-ttu-id="033b0-565">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-565">
         -DocumentEvents</span></span><br><span data-ttu-id="033b0-566">
         - File</span><span class="sxs-lookup"><span data-stu-id="033b0-566">
         - File</span></span><br><span data-ttu-id="033b0-567">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-567">
         -ImageCoercion</span></span><br><span data-ttu-id="033b0-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="033b0-568">
         -PdfFile</span></span><br><span data-ttu-id="033b0-569">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-569">
         - Selection</span></span><br><span data-ttu-id="033b0-570">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-570">
         - Settings</span></span><br><span data-ttu-id="033b0-571">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-571">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-572">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="033b0-572">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="033b0-573">- Contenu</span><span class="sxs-lookup"><span data-stu-id="033b0-573">- Content</span></span><br><span data-ttu-id="033b0-574">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-574">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="033b0-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="033b0-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="033b0-576">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="033b0-576">-ActiveView</span></span><br><span data-ttu-id="033b0-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-577">
         -CompressedFile</span></span><br><span data-ttu-id="033b0-578">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-578">
         -DocumentEvents</span></span><br><span data-ttu-id="033b0-579">
         - File</span><span class="sxs-lookup"><span data-stu-id="033b0-579">
         - File</span></span><br><span data-ttu-id="033b0-580">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-580">
         -ImageCoercion</span></span><br><span data-ttu-id="033b0-581">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="033b0-581">
         -PdfFile</span></span><br><span data-ttu-id="033b0-582">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-582">
         - Selection</span></span><br><span data-ttu-id="033b0-583">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-583">
         - Settings</span></span><br><span data-ttu-id="033b0-584">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-584">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-585">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="033b0-585">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="033b0-586">- Contenu</span><span class="sxs-lookup"><span data-stu-id="033b0-586">- Content</span></span><br><span data-ttu-id="033b0-587">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-587">
         - Taskpane</span></span><br><span data-ttu-id="033b0-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="033b0-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="033b0-590">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="033b0-590">-ActiveView</span></span><br><span data-ttu-id="033b0-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-591">
         -CompressedFile</span></span><br><span data-ttu-id="033b0-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-592">
         -DocumentEvents</span></span><br><span data-ttu-id="033b0-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="033b0-593">
         - File</span></span><br><span data-ttu-id="033b0-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-594">
         -ImageCoercion</span></span><br><span data-ttu-id="033b0-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="033b0-595">
         -PdfFile</span></span><br><span data-ttu-id="033b0-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-596">
         - Selection</span></span><br><span data-ttu-id="033b0-597">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-597">
         - Settings</span></span><br><span data-ttu-id="033b0-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-598">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-599">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="033b0-599">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="033b0-600">- Contenu</span><span class="sxs-lookup"><span data-stu-id="033b0-600">- Content</span></span><br><span data-ttu-id="033b0-601">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-601">
         - Taskpane</span></span><br><span data-ttu-id="033b0-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="033b0-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="033b0-604">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="033b0-604">-ActiveView</span></span><br><span data-ttu-id="033b0-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-605">
         -CompressedFile</span></span><br><span data-ttu-id="033b0-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-606">
         -DocumentEvents</span></span><br><span data-ttu-id="033b0-607">
         - File</span><span class="sxs-lookup"><span data-stu-id="033b0-607">
         - File</span></span><br><span data-ttu-id="033b0-608">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-608">
         -ImageCoercion</span></span><br><span data-ttu-id="033b0-609">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="033b0-609">
         -PdfFile</span></span><br><span data-ttu-id="033b0-610">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-610">
         - Selection</span></span><br><span data-ttu-id="033b0-611">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-611">
         - Settings</span></span><br><span data-ttu-id="033b0-612">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-612">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-613">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="033b0-613">Office for iOS</span></span></td>
    <td> <span data-ttu-id="033b0-614">- Contenu</span><span class="sxs-lookup"><span data-stu-id="033b0-614">- Content</span></span><br><span data-ttu-id="033b0-615">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-615">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="033b0-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="033b0-617">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="033b0-617">-ActiveView</span></span><br><span data-ttu-id="033b0-618">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-618">
         -CompressedFile</span></span><br><span data-ttu-id="033b0-619">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-619">
         -DocumentEvents</span></span><br><span data-ttu-id="033b0-620">
         - File</span><span class="sxs-lookup"><span data-stu-id="033b0-620">
         - File</span></span><br><span data-ttu-id="033b0-621">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="033b0-621">
         -PdfFile</span></span><br><span data-ttu-id="033b0-622">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-622">
         - Selection</span></span><br><span data-ttu-id="033b0-623">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-623">
         - Settings</span></span><br><span data-ttu-id="033b0-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-624">
         -TextCoercion</span></span><br><span data-ttu-id="033b0-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-625">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-626">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="033b0-626">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="033b0-627">- Contenu</span><span class="sxs-lookup"><span data-stu-id="033b0-627">- Content</span></span><br><span data-ttu-id="033b0-628">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-628">
         - Taskpane</span></span><br><span data-ttu-id="033b0-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="033b0-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="033b0-631">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="033b0-631">-ActiveView</span></span><br><span data-ttu-id="033b0-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-632">
         -CompressedFile</span></span><br><span data-ttu-id="033b0-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-633">
         -DocumentEvents</span></span><br><span data-ttu-id="033b0-634">
         - File</span><span class="sxs-lookup"><span data-stu-id="033b0-634">
         - File</span></span><br><span data-ttu-id="033b0-635">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-635">
         -ImageCoercion</span></span><br><span data-ttu-id="033b0-636">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="033b0-636">
         -PdfFile</span></span><br><span data-ttu-id="033b0-637">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-637">
         - Selection</span></span><br><span data-ttu-id="033b0-638">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-638">
         - Settings</span></span><br><span data-ttu-id="033b0-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-639">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-640">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="033b0-640">Office for Mac</span></span></td>
    <td> <span data-ttu-id="033b0-641">- Contenu</span><span class="sxs-lookup"><span data-stu-id="033b0-641">- Content</span></span><br><span data-ttu-id="033b0-642">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-642">
         - Taskpane</span></span><br><span data-ttu-id="033b0-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="033b0-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="033b0-645">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="033b0-645">-ActiveView</span></span><br><span data-ttu-id="033b0-646">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="033b0-646">
         -CompressedFile</span></span><br><span data-ttu-id="033b0-647">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-647">
         -DocumentEvents</span></span><br><span data-ttu-id="033b0-648">
         - File</span><span class="sxs-lookup"><span data-stu-id="033b0-648">
         - File</span></span><br><span data-ttu-id="033b0-649">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-649">
         -ImageCoercion</span></span><br><span data-ttu-id="033b0-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="033b0-650">
         -PdfFile</span></span><br><span data-ttu-id="033b0-651">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-651">
         - Selection</span></span><br><span data-ttu-id="033b0-652">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-652">
         - Settings</span></span><br><span data-ttu-id="033b0-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-653">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="033b0-654">OneNote</span><span class="sxs-lookup"><span data-stu-id="033b0-654">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="033b0-655">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="033b0-655">Platform</span></span></th>
    <th><span data-ttu-id="033b0-656">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="033b0-656">Extension points</span></span></th>
    <th><span data-ttu-id="033b0-657">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="033b0-657">API requirement sets</span></span></th>
    <th><span data-ttu-id="033b0-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="033b0-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-659">Office Online</span><span class="sxs-lookup"><span data-stu-id="033b0-659">Office Online</span></span></td>
    <td> <span data-ttu-id="033b0-660">- Contenu</span><span class="sxs-lookup"><span data-stu-id="033b0-660">- Content</span></span><br><span data-ttu-id="033b0-661">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-661">
         - Taskpane</span></span><br><span data-ttu-id="033b0-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="033b0-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="033b0-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="033b0-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="033b0-665">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="033b0-665">-DocumentEvents</span></span><br><span data-ttu-id="033b0-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-666">
         -HtmlCoercion</span></span><br><span data-ttu-id="033b0-667">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-667">
         -ImageCoercion</span></span><br><span data-ttu-id="033b0-668">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="033b0-668">
         - Settings</span></span><br><span data-ttu-id="033b0-669">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-669">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="033b0-670">Projet</span><span class="sxs-lookup"><span data-stu-id="033b0-670">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="033b0-671">Plateforme</span><span class="sxs-lookup"><span data-stu-id="033b0-671">Platform</span></span></th>
    <th><span data-ttu-id="033b0-672">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="033b0-672">Extension points</span></span></th>
    <th><span data-ttu-id="033b0-673">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="033b0-673">API requirement sets</span></span></th>
    <th><span data-ttu-id="033b0-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="033b0-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-675">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="033b0-675">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="033b0-676">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-676">- Taskpane</span></span></td>
    <td> <span data-ttu-id="033b0-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="033b0-678">- Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-678">- Selection</span></span><br><span data-ttu-id="033b0-679">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-679">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-680">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="033b0-680">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="033b0-681">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-681">- Taskpane</span></span></td>
    <td> <span data-ttu-id="033b0-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="033b0-683">- Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-683">- Selection</span></span><br><span data-ttu-id="033b0-684">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-684">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="033b0-685">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="033b0-685">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="033b0-686">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="033b0-686">- Taskpane</span></span></td>
    <td> <span data-ttu-id="033b0-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="033b0-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="033b0-688">- Selection</span><span class="sxs-lookup"><span data-stu-id="033b0-688">- Selection</span></span><br><span data-ttu-id="033b0-689">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="033b0-689">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="033b0-690">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="033b0-690">See also</span></span>

- [<span data-ttu-id="033b0-691">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="033b0-691">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="033b0-692">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="033b0-692">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="033b0-693">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="033b0-693">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="033b0-694">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="033b0-694">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
