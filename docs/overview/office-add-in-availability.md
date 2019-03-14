---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint, OneNote et Project.
ms.date: 03/07/2019
localization_priority: Priority
ms.openlocfilehash: 636c6290d8c67901beb195990593727485467460
ms.sourcegitcommit: 8e7b7b0cfb68b91a3a95585d094cf5f5ffd00178
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/09/2019
ms.locfileid: "30512880"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="e9d06-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="e9d06-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="e9d06-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="e9d06-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="e9d06-p102">Le numéro de build pour Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que les ensembles de conditions requises des API communes, ExcelApi 1.1 et WordApi 1.1.</span><span class="sxs-lookup"><span data-stu-id="e9d06-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="e9d06-108">Excel</span><span class="sxs-lookup"><span data-stu-id="e9d06-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="e9d06-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="e9d06-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="e9d06-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="e9d06-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="e9d06-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="e9d06-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="e9d06-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="e9d06-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="e9d06-113">Office Online</span></span></td>
    <td> <span data-ttu-id="e9d06-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-114">- TaskPane</span></span><br><span data-ttu-id="e9d06-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="e9d06-115">
        - Content</span></span><br><span data-ttu-id="e9d06-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="e9d06-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="e9d06-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e9d06-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e9d06-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e9d06-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e9d06-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e9d06-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e9d06-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e9d06-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e9d06-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e9d06-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-126">
        - BindingEvents</span></span><br><span data-ttu-id="e9d06-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-127">
        - CompressedFile</span></span><br><span data-ttu-id="e9d06-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-128">
        - DocumentEvents</span></span><br><span data-ttu-id="e9d06-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-129">
        - File</span></span><br><span data-ttu-id="e9d06-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-130">
        - MatrixBindings</span></span><br><span data-ttu-id="e9d06-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="e9d06-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-132">
        - Selection</span></span><br><span data-ttu-id="e9d06-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-133">
        - Settings</span></span><br><span data-ttu-id="e9d06-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-134">
        - TableBindings</span></span><br><span data-ttu-id="e9d06-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-135">
        - TableCoercion</span></span><br><span data-ttu-id="e9d06-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-136">
        - TextBindings</span></span><br><span data-ttu-id="e9d06-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-138">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="e9d06-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="e9d06-139">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-139">
        - TaskPane</span></span><br><span data-ttu-id="e9d06-140">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="e9d06-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="e9d06-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="e9d06-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="e9d06-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-142">
        - BindingEvents</span></span><br><span data-ttu-id="e9d06-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-143">
        - CompressedFile</span></span><br><span data-ttu-id="e9d06-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-144">
        - DocumentEvents</span></span><br><span data-ttu-id="e9d06-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-145">
        - File</span></span><br><span data-ttu-id="e9d06-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-146">
        - ImageCoercion</span></span><br><span data-ttu-id="e9d06-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-147">
        - MatrixBindings</span></span><br><span data-ttu-id="e9d06-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="e9d06-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-149">
        - Selection</span></span><br><span data-ttu-id="e9d06-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-150">
        - Settings</span></span><br><span data-ttu-id="e9d06-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-151">
        - TableBindings</span></span><br><span data-ttu-id="e9d06-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-152">
        - TableCoercion</span></span><br><span data-ttu-id="e9d06-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-153">
        - TextBindings</span></span><br><span data-ttu-id="e9d06-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-155">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="e9d06-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="e9d06-156">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-156">- TaskPane</span></span><br><span data-ttu-id="e9d06-157">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="e9d06-157">
        - Content</span></span></td>
    <td><span data-ttu-id="e9d06-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e9d06-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="e9d06-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="e9d06-160">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-160">- BindingEvents</span></span><br><span data-ttu-id="e9d06-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-161">
        - CompressedFile</span></span><br><span data-ttu-id="e9d06-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-162">
        - DocumentEvents</span></span><br><span data-ttu-id="e9d06-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-163">
        - File</span></span><br><span data-ttu-id="e9d06-164">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-164">
        - ImageCoercion</span></span><br><span data-ttu-id="e9d06-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-165">
        - MatrixBindings</span></span><br><span data-ttu-id="e9d06-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="e9d06-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-167">
        - Selection</span></span><br><span data-ttu-id="e9d06-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-168">
        - Settings</span></span><br><span data-ttu-id="e9d06-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-169">
        - TableBindings</span></span><br><span data-ttu-id="e9d06-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-170">
        - TableCoercion</span></span><br><span data-ttu-id="e9d06-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-171">
        - TextBindings</span></span><br><span data-ttu-id="e9d06-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-173">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="e9d06-173">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="e9d06-174">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-174">- TaskPane</span></span><br><span data-ttu-id="e9d06-175">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="e9d06-175">
        - Content</span></span><br><span data-ttu-id="e9d06-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e9d06-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e9d06-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e9d06-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e9d06-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e9d06-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e9d06-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e9d06-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e9d06-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e9d06-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e9d06-186">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-186">- BindingEvents</span></span><br><span data-ttu-id="e9d06-187">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-187">
        - CompressedFile</span></span><br><span data-ttu-id="e9d06-188">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-188">
        - DocumentEvents</span></span><br><span data-ttu-id="e9d06-189">
        - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-189">
        - File</span></span><br><span data-ttu-id="e9d06-190">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-190">
        - ImageCoercion</span></span><br><span data-ttu-id="e9d06-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-191">
        - MatrixBindings</span></span><br><span data-ttu-id="e9d06-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="e9d06-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-193">
        - Selection</span></span><br><span data-ttu-id="e9d06-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-194">
        - Settings</span></span><br><span data-ttu-id="e9d06-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-195">
        - TableBindings</span></span><br><span data-ttu-id="e9d06-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-196">
        - TableCoercion</span></span><br><span data-ttu-id="e9d06-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-197">
        - TextBindings</span></span><br><span data-ttu-id="e9d06-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-199">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="e9d06-199">Office for iPad</span></span></td>
    <td><span data-ttu-id="e9d06-200">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-200">- TaskPane</span></span><br><span data-ttu-id="e9d06-201">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="e9d06-201">
        - Content</span></span></td>
    <td><span data-ttu-id="e9d06-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e9d06-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e9d06-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e9d06-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e9d06-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e9d06-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e9d06-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e9d06-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e9d06-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e9d06-211">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-211">- BindingEvents</span></span><br><span data-ttu-id="e9d06-212">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-212">
        - CompressedFile</span></span><br><span data-ttu-id="e9d06-213">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-213">
        - DocumentEvents</span></span><br><span data-ttu-id="e9d06-214">
        - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-214">
        - File</span></span><br><span data-ttu-id="e9d06-215">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-215">
        - ImageCoercion</span></span><br><span data-ttu-id="e9d06-216">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-216">
        - MatrixBindings</span></span><br><span data-ttu-id="e9d06-217">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-217">
        - MatrixCoercion</span></span><br><span data-ttu-id="e9d06-218">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-218">
        - Selection</span></span><br><span data-ttu-id="e9d06-219">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-219">
        - Settings</span></span><br><span data-ttu-id="e9d06-220">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-220">
        - TableBindings</span></span><br><span data-ttu-id="e9d06-221">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-221">
        - TableCoercion</span></span><br><span data-ttu-id="e9d06-222">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-222">
        - TextBindings</span></span><br><span data-ttu-id="e9d06-223">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-223">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-224">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="e9d06-224">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="e9d06-225">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-225">- TaskPane</span></span><br><span data-ttu-id="e9d06-226">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="e9d06-226">
        - Content</span></span></td>
    <td><span data-ttu-id="e9d06-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e9d06-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="e9d06-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="e9d06-229">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-229">- BindingEvents</span></span><br><span data-ttu-id="e9d06-230">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-230">
        - CompressedFile</span></span><br><span data-ttu-id="e9d06-231">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-231">
        - DocumentEvents</span></span><br><span data-ttu-id="e9d06-232">
        - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-232">
        - File</span></span><br><span data-ttu-id="e9d06-233">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-233">
        - ImageCoercion</span></span><br><span data-ttu-id="e9d06-234">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-234">
        - MatrixBindings</span></span><br><span data-ttu-id="e9d06-235">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-235">
        - MatrixCoercion</span></span><br><span data-ttu-id="e9d06-236">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-236">
        - PdfFile</span></span><br><span data-ttu-id="e9d06-237">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-237">
        - Selection</span></span><br><span data-ttu-id="e9d06-238">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-238">
        - Settings</span></span><br><span data-ttu-id="e9d06-239">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-239">
        - TableBindings</span></span><br><span data-ttu-id="e9d06-240">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-240">
        - TableCoercion</span></span><br><span data-ttu-id="e9d06-241">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-241">
        - TextBindings</span></span><br><span data-ttu-id="e9d06-242">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-242">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-243">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="e9d06-243">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="e9d06-244">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-244">- TaskPane</span></span><br><span data-ttu-id="e9d06-245">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="e9d06-245">
        - Content</span></span><br><span data-ttu-id="e9d06-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e9d06-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e9d06-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e9d06-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e9d06-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e9d06-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e9d06-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e9d06-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e9d06-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e9d06-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e9d06-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-256">- BindingEvents</span></span><br><span data-ttu-id="e9d06-257">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-257">
        - CompressedFile</span></span><br><span data-ttu-id="e9d06-258">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-258">
        - DocumentEvents</span></span><br><span data-ttu-id="e9d06-259">
        - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-259">
        - File</span></span><br><span data-ttu-id="e9d06-260">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-260">
        - ImageCoercion</span></span><br><span data-ttu-id="e9d06-261">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-261">
        - MatrixBindings</span></span><br><span data-ttu-id="e9d06-262">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-262">
        - MatrixCoercion</span></span><br><span data-ttu-id="e9d06-263">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-263">
        - PdfFile</span></span><br><span data-ttu-id="e9d06-264">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-264">
        - Selection</span></span><br><span data-ttu-id="e9d06-265">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-265">
        - Settings</span></span><br><span data-ttu-id="e9d06-266">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-266">
        - TableBindings</span></span><br><span data-ttu-id="e9d06-267">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-267">
        - TableCoercion</span></span><br><span data-ttu-id="e9d06-268">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-268">
        - TextBindings</span></span><br><span data-ttu-id="e9d06-269">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-269">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="e9d06-270">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="e9d06-270">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="e9d06-271">Outlook</span><span class="sxs-lookup"><span data-stu-id="e9d06-271">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e9d06-272">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="e9d06-272">Platform</span></span></th>
    <th><span data-ttu-id="e9d06-273">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="e9d06-273">Extension points</span></span></th>
    <th><span data-ttu-id="e9d06-274">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="e9d06-274">API requirement sets</span></span></th>
    <th><span data-ttu-id="e9d06-275"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="e9d06-275"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-276">Office Online</span><span class="sxs-lookup"><span data-stu-id="e9d06-276">Office Online</span></span></td>
    <td> <span data-ttu-id="e9d06-277">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="e9d06-277">- Mail Read</span></span><br><span data-ttu-id="e9d06-278">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="e9d06-278">
      - Mail Compose</span></span><br><span data-ttu-id="e9d06-279">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-279">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9d06-280">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-280">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e9d06-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e9d06-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e9d06-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e9d06-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e9d06-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e9d06-286">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-286">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e9d06-287">Non disponible</span><span class="sxs-lookup"><span data-stu-id="e9d06-287">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-288">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="e9d06-288">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e9d06-289">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="e9d06-289">- Mail Read</span></span><br><span data-ttu-id="e9d06-290">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="e9d06-290">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="e9d06-291">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-291">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e9d06-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e9d06-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e9d06-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="e9d06-295">Non disponible</span><span class="sxs-lookup"><span data-stu-id="e9d06-295">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-296">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="e9d06-296">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e9d06-297">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="e9d06-297">- Mail Read</span></span><br><span data-ttu-id="e9d06-298">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="e9d06-298">
      - Mail Compose</span></span><br><span data-ttu-id="e9d06-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e9d06-300">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="e9d06-300">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e9d06-301">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-301">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e9d06-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e9d06-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e9d06-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e9d06-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e9d06-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e9d06-307">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-307">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e9d06-308">Non disponible</span><span class="sxs-lookup"><span data-stu-id="e9d06-308">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-309">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="e9d06-309">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="e9d06-310">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="e9d06-310">- Mail Read</span></span><br><span data-ttu-id="e9d06-311">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="e9d06-311">
      - Mail Compose</span></span><br><span data-ttu-id="e9d06-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e9d06-313">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="e9d06-313">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e9d06-314">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-314">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e9d06-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e9d06-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e9d06-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e9d06-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e9d06-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e9d06-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e9d06-321">Non disponible</span><span class="sxs-lookup"><span data-stu-id="e9d06-321">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-322">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="e9d06-322">Office for iOS</span></span></td>
    <td> <span data-ttu-id="e9d06-323">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="e9d06-323">- Mail Read</span></span><br><span data-ttu-id="e9d06-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9d06-325">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-325">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e9d06-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e9d06-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e9d06-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e9d06-329">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-329">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e9d06-330">Non disponible</span><span class="sxs-lookup"><span data-stu-id="e9d06-330">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-331">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="e9d06-331">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e9d06-332">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="e9d06-332">- Mail Read</span></span><br><span data-ttu-id="e9d06-333">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="e9d06-333">
      - Mail Compose</span></span><br><span data-ttu-id="e9d06-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9d06-335">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-335">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e9d06-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e9d06-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e9d06-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e9d06-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e9d06-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e9d06-341">Non disponible</span><span class="sxs-lookup"><span data-stu-id="e9d06-341">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-342">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="e9d06-342">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="e9d06-343">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="e9d06-343">- Mail Read</span></span><br><span data-ttu-id="e9d06-344">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="e9d06-344">
      - Mail Compose</span></span><br><span data-ttu-id="e9d06-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9d06-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e9d06-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e9d06-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e9d06-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e9d06-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e9d06-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e9d06-352">Non disponible</span><span class="sxs-lookup"><span data-stu-id="e9d06-352">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-353">Office pour Android</span><span class="sxs-lookup"><span data-stu-id="e9d06-353">Office for Android</span></span></td>
    <td> <span data-ttu-id="e9d06-354">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="e9d06-354">- Mail Read</span></span><br><span data-ttu-id="e9d06-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9d06-356">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-356">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e9d06-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e9d06-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e9d06-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e9d06-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e9d06-361">Non disponible</span><span class="sxs-lookup"><span data-stu-id="e9d06-361">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="e9d06-362">Word</span><span class="sxs-lookup"><span data-stu-id="e9d06-362">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e9d06-363">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="e9d06-363">Platform</span></span></th>
    <th><span data-ttu-id="e9d06-364">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="e9d06-364">Extension points</span></span></th>
    <th><span data-ttu-id="e9d06-365">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="e9d06-365">API requirement sets</span></span></th>
    <th><span data-ttu-id="e9d06-366"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="e9d06-366"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-367">Office Online</span><span class="sxs-lookup"><span data-stu-id="e9d06-367">Office Online</span></span></td>
    <td> <span data-ttu-id="e9d06-368">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-368">- TaskPane</span></span><br><span data-ttu-id="e9d06-369">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-369">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9d06-370">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-370">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e9d06-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e9d06-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e9d06-373">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-373">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e9d06-374">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-374">- BindingEvents</span></span><br><span data-ttu-id="e9d06-375">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e9d06-375">
         - CustomXmlParts</span></span><br><span data-ttu-id="e9d06-376">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-376">
         - DocumentEvents</span></span><br><span data-ttu-id="e9d06-377">
         - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-377">
         - File</span></span><br><span data-ttu-id="e9d06-378">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-378">
         - HtmlCoercion</span></span><br><span data-ttu-id="e9d06-379">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-379">
         - ImageCoercion</span></span><br><span data-ttu-id="e9d06-380">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-380">
         - MatrixBindings</span></span><br><span data-ttu-id="e9d06-381">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-381">
         - MatrixCoercion</span></span><br><span data-ttu-id="e9d06-382">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-382">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e9d06-383">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-383">
         - PdfFile</span></span><br><span data-ttu-id="e9d06-384">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-384">
         - Selection</span></span><br><span data-ttu-id="e9d06-385">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-385">
         - Settings</span></span><br><span data-ttu-id="e9d06-386">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-386">
         - TableBindings</span></span><br><span data-ttu-id="e9d06-387">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-387">
         - TableCoercion</span></span><br><span data-ttu-id="e9d06-388">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-388">
         - TextBindings</span></span><br><span data-ttu-id="e9d06-389">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-389">
         - TextCoercion</span></span><br><span data-ttu-id="e9d06-390">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-390">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-391">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="e9d06-391">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e9d06-392">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-392">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e9d06-393">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="e9d06-393">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="e9d06-394">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-394">- BindingEvents</span></span><br><span data-ttu-id="e9d06-395">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-395">
         - CompressedFile</span></span><br><span data-ttu-id="e9d06-396">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e9d06-396">
         - CustomXmlParts</span></span><br><span data-ttu-id="e9d06-397">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-397">
         - DocumentEvents</span></span><br><span data-ttu-id="e9d06-398">
         - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-398">
         - File</span></span><br><span data-ttu-id="e9d06-399">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-399">
         - HtmlCoercion</span></span><br><span data-ttu-id="e9d06-400">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-400">
         - ImageCoercion</span></span><br><span data-ttu-id="e9d06-401">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-401">
         - MatrixBindings</span></span><br><span data-ttu-id="e9d06-402">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-402">
         - MatrixCoercion</span></span><br><span data-ttu-id="e9d06-403">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-403">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e9d06-404">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-404">
         - PdfFile</span></span><br><span data-ttu-id="e9d06-405">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-405">
         - Selection</span></span><br><span data-ttu-id="e9d06-406">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-406">
         - Settings</span></span><br><span data-ttu-id="e9d06-407">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-407">
         - TableBindings</span></span><br><span data-ttu-id="e9d06-408">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-408">
         - TableCoercion</span></span><br><span data-ttu-id="e9d06-409">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-409">
         - TextBindings</span></span><br><span data-ttu-id="e9d06-410">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-410">
         - TextCoercion</span></span><br><span data-ttu-id="e9d06-411">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-411">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-412">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="e9d06-412">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e9d06-413">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-413">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e9d06-414">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-414">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e9d06-415">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="e9d06-415">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="e9d06-416">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-416">- BindingEvents</span></span><br><span data-ttu-id="e9d06-417">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-417">
         - CompressedFile</span></span><br><span data-ttu-id="e9d06-418">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e9d06-418">
         - CustomXmlParts</span></span><br><span data-ttu-id="e9d06-419">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-419">
         - DocumentEvents</span></span><br><span data-ttu-id="e9d06-420">
         - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-420">
         - File</span></span><br><span data-ttu-id="e9d06-421">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-421">
         - HtmlCoercion</span></span><br><span data-ttu-id="e9d06-422">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-422">
         - ImageCoercion</span></span><br><span data-ttu-id="e9d06-423">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-423">
         - MatrixBindings</span></span><br><span data-ttu-id="e9d06-424">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-424">
         - MatrixCoercion</span></span><br><span data-ttu-id="e9d06-425">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-425">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e9d06-426">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-426">
         - PdfFile</span></span><br><span data-ttu-id="e9d06-427">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-427">
         - Selection</span></span><br><span data-ttu-id="e9d06-428">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-428">
         - Settings</span></span><br><span data-ttu-id="e9d06-429">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-429">
         - TableBindings</span></span><br><span data-ttu-id="e9d06-430">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-430">
         - TableCoercion</span></span><br><span data-ttu-id="e9d06-431">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-431">
         - TextBindings</span></span><br><span data-ttu-id="e9d06-432">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-432">
         - TextCoercion</span></span><br><span data-ttu-id="e9d06-433">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-433">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-434">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="e9d06-434">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="e9d06-435">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-435">- TaskPane</span></span><br><span data-ttu-id="e9d06-436">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-436">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9d06-437">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-437">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e9d06-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e9d06-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e9d06-440">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-440">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e9d06-441">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-441">- BindingEvents</span></span><br><span data-ttu-id="e9d06-442">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-442">
         - CompressedFile</span></span><br><span data-ttu-id="e9d06-443">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e9d06-443">
         - CustomXmlParts</span></span><br><span data-ttu-id="e9d06-444">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-444">
         - DocumentEvents</span></span><br><span data-ttu-id="e9d06-445">
         - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-445">
         - File</span></span><br><span data-ttu-id="e9d06-446">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-446">
         - HtmlCoercion</span></span><br><span data-ttu-id="e9d06-447">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-447">
         - ImageCoercion</span></span><br><span data-ttu-id="e9d06-448">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-448">
         - MatrixBindings</span></span><br><span data-ttu-id="e9d06-449">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-449">
         - MatrixCoercion</span></span><br><span data-ttu-id="e9d06-450">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-450">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e9d06-451">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-451">
         - PdfFile</span></span><br><span data-ttu-id="e9d06-452">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-452">
         - Selection</span></span><br><span data-ttu-id="e9d06-453">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-453">
         - Settings</span></span><br><span data-ttu-id="e9d06-454">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-454">
         - TableBindings</span></span><br><span data-ttu-id="e9d06-455">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-455">
         - TableCoercion</span></span><br><span data-ttu-id="e9d06-456">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-456">
         - TextBindings</span></span><br><span data-ttu-id="e9d06-457">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-457">
         - TextCoercion</span></span><br><span data-ttu-id="e9d06-458">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-458">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-459">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="e9d06-459">Office for iPad</span></span></td>
    <td> <span data-ttu-id="e9d06-460">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-460">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e9d06-461">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-461">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e9d06-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e9d06-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e9d06-464">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e9d06-464">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e9d06-465">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-465">- BindingEvents</span></span><br><span data-ttu-id="e9d06-466">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-466">
         - CompressedFile</span></span><br><span data-ttu-id="e9d06-467">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e9d06-467">
         - CustomXmlParts</span></span><br><span data-ttu-id="e9d06-468">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-468">
         - DocumentEvents</span></span><br><span data-ttu-id="e9d06-469">
         - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-469">
         - File</span></span><br><span data-ttu-id="e9d06-470">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-470">
         - HtmlCoercion</span></span><br><span data-ttu-id="e9d06-471">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-471">
         - ImageCoercion</span></span><br><span data-ttu-id="e9d06-472">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-472">
         - MatrixBindings</span></span><br><span data-ttu-id="e9d06-473">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-473">
         - MatrixCoercion</span></span><br><span data-ttu-id="e9d06-474">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-474">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e9d06-475">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-475">
         - PdfFile</span></span><br><span data-ttu-id="e9d06-476">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-476">
         - Selection</span></span><br><span data-ttu-id="e9d06-477">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-477">
         - Settings</span></span><br><span data-ttu-id="e9d06-478">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-478">
         - TableBindings</span></span><br><span data-ttu-id="e9d06-479">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-479">
         - TableCoercion</span></span><br><span data-ttu-id="e9d06-480">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-480">
         - TextBindings</span></span><br><span data-ttu-id="e9d06-481">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-481">
         - TextCoercion</span></span><br><span data-ttu-id="e9d06-482">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-482">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-483">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="e9d06-483">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e9d06-484">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-484">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e9d06-485">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-485">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e9d06-486">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="e9d06-486">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="e9d06-487">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-487">- BindingEvents</span></span><br><span data-ttu-id="e9d06-488">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-488">
         - CompressedFile</span></span><br><span data-ttu-id="e9d06-489">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e9d06-489">
         - CustomXmlParts</span></span><br><span data-ttu-id="e9d06-490">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-490">
         - DocumentEvents</span></span><br><span data-ttu-id="e9d06-491">
         - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-491">
         - File</span></span><br><span data-ttu-id="e9d06-492">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-492">
         - HtmlCoercion</span></span><br><span data-ttu-id="e9d06-493">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-493">
         - ImageCoercion</span></span><br><span data-ttu-id="e9d06-494">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-494">
         - MatrixBindings</span></span><br><span data-ttu-id="e9d06-495">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-495">
         - MatrixCoercion</span></span><br><span data-ttu-id="e9d06-496">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-496">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e9d06-497">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-497">
         - PdfFile</span></span><br><span data-ttu-id="e9d06-498">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-498">
         - Selection</span></span><br><span data-ttu-id="e9d06-499">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-499">
         - Settings</span></span><br><span data-ttu-id="e9d06-500">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-500">
         - TableBindings</span></span><br><span data-ttu-id="e9d06-501">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-501">
         - TableCoercion</span></span><br><span data-ttu-id="e9d06-502">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-502">
         - TextBindings</span></span><br><span data-ttu-id="e9d06-503">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-503">
         - TextCoercion</span></span><br><span data-ttu-id="e9d06-504">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-504">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-505">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="e9d06-505">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="e9d06-506">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-506">- TaskPane</span></span><br><span data-ttu-id="e9d06-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9d06-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e9d06-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e9d06-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e9d06-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e9d06-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e9d06-512">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-512">- BindingEvents</span></span><br><span data-ttu-id="e9d06-513">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-513">
         - CompressedFile</span></span><br><span data-ttu-id="e9d06-514">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e9d06-514">
         - CustomXmlParts</span></span><br><span data-ttu-id="e9d06-515">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-515">
         - DocumentEvents</span></span><br><span data-ttu-id="e9d06-516">
         - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-516">
         - File</span></span><br><span data-ttu-id="e9d06-517">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-517">
         - HtmlCoercion</span></span><br><span data-ttu-id="e9d06-518">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-518">
         - ImageCoercion</span></span><br><span data-ttu-id="e9d06-519">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-519">
         - MatrixBindings</span></span><br><span data-ttu-id="e9d06-520">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-520">
         - MatrixCoercion</span></span><br><span data-ttu-id="e9d06-521">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-521">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e9d06-522">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-522">
         - PdfFile</span></span><br><span data-ttu-id="e9d06-523">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-523">
         - Selection</span></span><br><span data-ttu-id="e9d06-524">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-524">
         - Settings</span></span><br><span data-ttu-id="e9d06-525">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-525">
         - TableBindings</span></span><br><span data-ttu-id="e9d06-526">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-526">
         - TableCoercion</span></span><br><span data-ttu-id="e9d06-527">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9d06-527">
         - TextBindings</span></span><br><span data-ttu-id="e9d06-528">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-528">
         - TextCoercion</span></span><br><span data-ttu-id="e9d06-529">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-529">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="e9d06-530">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="e9d06-530">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="e9d06-531">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="e9d06-531">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e9d06-532">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="e9d06-532">Platform</span></span></th>
    <th><span data-ttu-id="e9d06-533">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="e9d06-533">Extension points</span></span></th>
    <th><span data-ttu-id="e9d06-534">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="e9d06-534">API requirement sets</span></span></th>
    <th><span data-ttu-id="e9d06-535"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="e9d06-535"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-536">Office Online</span><span class="sxs-lookup"><span data-stu-id="e9d06-536">Office Online</span></span></td>
    <td> <span data-ttu-id="e9d06-537">- Contenu</span><span class="sxs-lookup"><span data-stu-id="e9d06-537">- Content</span></span><br><span data-ttu-id="e9d06-538">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-538">
         - TaskPane</span></span><br><span data-ttu-id="e9d06-539">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-539">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9d06-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e9d06-541">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e9d06-541">- ActiveView</span></span><br><span data-ttu-id="e9d06-542">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-542">
         - CompressedFile</span></span><br><span data-ttu-id="e9d06-543">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-543">
         - DocumentEvents</span></span><br><span data-ttu-id="e9d06-544">
         - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-544">
         - File</span></span><br><span data-ttu-id="e9d06-545">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-545">
         - ImageCoercion</span></span><br><span data-ttu-id="e9d06-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-546">
         - PdfFile</span></span><br><span data-ttu-id="e9d06-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-547">
         - Selection</span></span><br><span data-ttu-id="e9d06-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-548">
         - Settings</span></span><br><span data-ttu-id="e9d06-549">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-549">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-550">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="e9d06-550">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e9d06-551">- Contenu</span><span class="sxs-lookup"><span data-stu-id="e9d06-551">- Content</span></span><br><span data-ttu-id="e9d06-552">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-552">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="e9d06-553">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="e9d06-553">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="e9d06-554">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e9d06-554">- ActiveView</span></span><br><span data-ttu-id="e9d06-555">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-555">
         - CompressedFile</span></span><br><span data-ttu-id="e9d06-556">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-556">
         - DocumentEvents</span></span><br><span data-ttu-id="e9d06-557">
         - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-557">
         - File</span></span><br><span data-ttu-id="e9d06-558">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-558">
         - ImageCoercion</span></span><br><span data-ttu-id="e9d06-559">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-559">
         - PdfFile</span></span><br><span data-ttu-id="e9d06-560">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-560">
         - Selection</span></span><br><span data-ttu-id="e9d06-561">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-561">
         - Settings</span></span><br><span data-ttu-id="e9d06-562">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-562">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-563">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="e9d06-563">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e9d06-564">- Contenu</span><span class="sxs-lookup"><span data-stu-id="e9d06-564">- Content</span></span><br><span data-ttu-id="e9d06-565">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-565">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="e9d06-566">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="e9d06-566">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="e9d06-567">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e9d06-567">- ActiveView</span></span><br><span data-ttu-id="e9d06-568">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-568">
         - CompressedFile</span></span><br><span data-ttu-id="e9d06-569">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-569">
         - DocumentEvents</span></span><br><span data-ttu-id="e9d06-570">
         - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-570">
         - File</span></span><br><span data-ttu-id="e9d06-571">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-571">
         - ImageCoercion</span></span><br><span data-ttu-id="e9d06-572">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-572">
         - PdfFile</span></span><br><span data-ttu-id="e9d06-573">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-573">
         - Selection</span></span><br><span data-ttu-id="e9d06-574">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-574">
         - Settings</span></span><br><span data-ttu-id="e9d06-575">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-575">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-576">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="e9d06-576">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="e9d06-577">- Contenu</span><span class="sxs-lookup"><span data-stu-id="e9d06-577">- Content</span></span><br><span data-ttu-id="e9d06-578">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-578">
         - TaskPane</span></span><br><span data-ttu-id="e9d06-579">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-579">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9d06-580">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-580">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e9d06-581">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e9d06-581">- ActiveView</span></span><br><span data-ttu-id="e9d06-582">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-582">
         - CompressedFile</span></span><br><span data-ttu-id="e9d06-583">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-583">
         - DocumentEvents</span></span><br><span data-ttu-id="e9d06-584">
         - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-584">
         - File</span></span><br><span data-ttu-id="e9d06-585">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-585">
         - ImageCoercion</span></span><br><span data-ttu-id="e9d06-586">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-586">
         - PdfFile</span></span><br><span data-ttu-id="e9d06-587">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-587">
         - Selection</span></span><br><span data-ttu-id="e9d06-588">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-588">
         - Settings</span></span><br><span data-ttu-id="e9d06-589">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-589">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-590">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="e9d06-590">Office for iPad</span></span></td>
    <td> <span data-ttu-id="e9d06-591">- Contenu</span><span class="sxs-lookup"><span data-stu-id="e9d06-591">- Content</span></span><br><span data-ttu-id="e9d06-592">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-592">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="e9d06-593">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-593">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="e9d06-594">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e9d06-594">- ActiveView</span></span><br><span data-ttu-id="e9d06-595">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-595">
         - CompressedFile</span></span><br><span data-ttu-id="e9d06-596">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-596">
         - DocumentEvents</span></span><br><span data-ttu-id="e9d06-597">
         - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-597">
         - File</span></span><br><span data-ttu-id="e9d06-598">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-598">
         - PdfFile</span></span><br><span data-ttu-id="e9d06-599">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-599">
         - Selection</span></span><br><span data-ttu-id="e9d06-600">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-600">
         - Settings</span></span><br><span data-ttu-id="e9d06-601">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-601">
         - TextCoercion</span></span><br><span data-ttu-id="e9d06-602">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-602">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-603">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="e9d06-603">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e9d06-604">- Contenu</span><span class="sxs-lookup"><span data-stu-id="e9d06-604">- Content</span></span><br><span data-ttu-id="e9d06-605">Volet Office 
         -/td></span><span class="sxs-lookup"><span data-stu-id="e9d06-605">
         - TaskPane/td></span></span> <td> <span data-ttu-id="e9d06-606">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="e9d06-606">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="e9d06-607">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e9d06-607">- ActiveView</span></span><br><span data-ttu-id="e9d06-608">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-608">
         - CompressedFile</span></span><br><span data-ttu-id="e9d06-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-609">
         - DocumentEvents</span></span><br><span data-ttu-id="e9d06-610">
         - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-610">
         - File</span></span><br><span data-ttu-id="e9d06-611">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-611">
         - ImageCoercion</span></span><br><span data-ttu-id="e9d06-612">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-612">
         - PdfFile</span></span><br><span data-ttu-id="e9d06-613">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-613">
         - Selection</span></span><br><span data-ttu-id="e9d06-614">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-614">
         - Settings</span></span><br><span data-ttu-id="e9d06-615">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-615">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-616">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="e9d06-616">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="e9d06-617">- Contenu</span><span class="sxs-lookup"><span data-stu-id="e9d06-617">- Content</span></span><br><span data-ttu-id="e9d06-618">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-618">
         - TaskPane</span></span><br><span data-ttu-id="e9d06-619">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-619">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9d06-620">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-620">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e9d06-621">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e9d06-621">- ActiveView</span></span><br><span data-ttu-id="e9d06-622">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-622">
         - CompressedFile</span></span><br><span data-ttu-id="e9d06-623">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-623">
         - DocumentEvents</span></span><br><span data-ttu-id="e9d06-624">
         - File</span><span class="sxs-lookup"><span data-stu-id="e9d06-624">
         - File</span></span><br><span data-ttu-id="e9d06-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-625">
         - ImageCoercion</span></span><br><span data-ttu-id="e9d06-626">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9d06-626">
         - PdfFile</span></span><br><span data-ttu-id="e9d06-627">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-627">
         - Selection</span></span><br><span data-ttu-id="e9d06-628">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-628">
         - Settings</span></span><br><span data-ttu-id="e9d06-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-629">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="e9d06-630">*&ast; : ajouté avec les mises à jour après la publication.*</span><span class="sxs-lookup"><span data-stu-id="e9d06-630">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="e9d06-631">OneNote</span><span class="sxs-lookup"><span data-stu-id="e9d06-631">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e9d06-632">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="e9d06-632">Platform</span></span></th>
    <th><span data-ttu-id="e9d06-633">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="e9d06-633">Extension points</span></span></th>
    <th><span data-ttu-id="e9d06-634">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="e9d06-634">API requirement sets</span></span></th>
    <th><span data-ttu-id="e9d06-635"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="e9d06-635"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-636">Office Online</span><span class="sxs-lookup"><span data-stu-id="e9d06-636">Office Online</span></span></td>
    <td> <span data-ttu-id="e9d06-637">- Contenu</span><span class="sxs-lookup"><span data-stu-id="e9d06-637">- Content</span></span><br><span data-ttu-id="e9d06-638">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-638">
         - TaskPane</span></span><br><span data-ttu-id="e9d06-639">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-639">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9d06-640">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-640">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="e9d06-641">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-641">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e9d06-642">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9d06-642">- DocumentEvents</span></span><br><span data-ttu-id="e9d06-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="e9d06-644">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-644">
         - ImageCoercion</span></span><br><span data-ttu-id="e9d06-645">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e9d06-645">
         - Settings</span></span><br><span data-ttu-id="e9d06-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-646">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="e9d06-647">Projet</span><span class="sxs-lookup"><span data-stu-id="e9d06-647">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e9d06-648">Plateforme</span><span class="sxs-lookup"><span data-stu-id="e9d06-648">Platform</span></span></th>
    <th><span data-ttu-id="e9d06-649">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="e9d06-649">Extension points</span></span></th>
    <th><span data-ttu-id="e9d06-650">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="e9d06-650">API requirement sets</span></span></th>
    <th><span data-ttu-id="e9d06-651"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="e9d06-651"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-652">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="e9d06-652">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e9d06-653">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-653">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e9d06-654">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-654">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e9d06-655">- Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-655">- Selection</span></span><br><span data-ttu-id="e9d06-656">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-656">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-657">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="e9d06-657">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e9d06-658">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-658">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e9d06-659">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-659">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e9d06-660">- Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-660">- Selection</span></span><br><span data-ttu-id="e9d06-661">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-661">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9d06-662">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="e9d06-662">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="e9d06-663">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-663">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e9d06-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9d06-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e9d06-665">- Selection</span><span class="sxs-lookup"><span data-stu-id="e9d06-665">- Selection</span></span><br><span data-ttu-id="e9d06-666">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9d06-666">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="e9d06-667">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e9d06-667">See also</span></span>

- [<span data-ttu-id="e9d06-668">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-668">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="e9d06-669">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="e9d06-669">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="e9d06-670">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="e9d06-670">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="e9d06-671">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="e9d06-671">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
