---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint, OneNote et Project.
ms.date: 02/20/2019
localization_priority: Priority
ms.openlocfilehash: a3e9c508a5bae0e7eb660458835b9242d0602818
ms.sourcegitcommit: 8e20e7663be2aaa0f7a5436a965324d171bc667d
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/28/2019
ms.locfileid: "30199612"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="fd4a0-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="fd4a0-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="fd4a0-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="fd4a0-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="fd4a0-p102">Le numéro de build pour Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que les ensembles de conditions requises des API communes, ExcelApi 1.1 et WordApi 1.1.</span><span class="sxs-lookup"><span data-stu-id="fd4a0-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="fd4a0-108">Excel</span><span class="sxs-lookup"><span data-stu-id="fd4a0-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="fd4a0-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="fd4a0-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="fd4a0-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="fd4a0-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="fd4a0-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="fd4a0-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="fd4a0-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="fd4a0-113">Office Online</span></span></td>
    <td> <span data-ttu-id="fd4a0-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-114">- TaskPane</span></span><br><span data-ttu-id="fd4a0-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="fd4a0-115">
        - Content</span></span><br><span data-ttu-id="fd4a0-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="fd4a0-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="fd4a0-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fd4a0-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="fd4a0-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="fd4a0-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="fd4a0-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="fd4a0-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="fd4a0-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="fd4a0-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="fd4a0-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="fd4a0-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-126">
        - BindingEvents</span></span><br><span data-ttu-id="fd4a0-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-127">
        - CompressedFile</span></span><br><span data-ttu-id="fd4a0-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-128">
        - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-129">
        - File</span></span><br><span data-ttu-id="fd4a0-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-130">
        - MatrixBindings</span></span><br><span data-ttu-id="fd4a0-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="fd4a0-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-132">
        - Selection</span></span><br><span data-ttu-id="fd4a0-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-133">
        - Settings</span></span><br><span data-ttu-id="fd4a0-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-134">
        - TableBindings</span></span><br><span data-ttu-id="fd4a0-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-135">
        - TableCoercion</span></span><br><span data-ttu-id="fd4a0-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-136">
        - TextBindings</span></span><br><span data-ttu-id="fd4a0-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-138">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="fd4a0-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="fd4a0-139">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-139">
        - TaskPane</span></span><br><span data-ttu-id="fd4a0-140">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="fd4a0-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="fd4a0-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="fd4a0-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="fd4a0-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-142">
        - BindingEvents</span></span><br><span data-ttu-id="fd4a0-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-143">
        - CompressedFile</span></span><br><span data-ttu-id="fd4a0-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-144">
        - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-145">
        - File</span></span><br><span data-ttu-id="fd4a0-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-146">
        - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-147">
        - MatrixBindings</span></span><br><span data-ttu-id="fd4a0-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="fd4a0-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-149">
        - Selection</span></span><br><span data-ttu-id="fd4a0-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-150">
        - Settings</span></span><br><span data-ttu-id="fd4a0-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-151">
        - TableBindings</span></span><br><span data-ttu-id="fd4a0-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-152">
        - TableCoercion</span></span><br><span data-ttu-id="fd4a0-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-153">
        - TextBindings</span></span><br><span data-ttu-id="fd4a0-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-155">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="fd4a0-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="fd4a0-156">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-156">- TaskPane</span></span><br><span data-ttu-id="fd4a0-157">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="fd4a0-157">
        - Content</span></span></td>
    <td><span data-ttu-id="fd4a0-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fd4a0-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="fd4a0-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="fd4a0-160">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-160">- BindingEvents</span></span><br><span data-ttu-id="fd4a0-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-161">
        - CompressedFile</span></span><br><span data-ttu-id="fd4a0-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-162">
        - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-163">
        - File</span></span><br><span data-ttu-id="fd4a0-164">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-164">
        - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-165">
        - MatrixBindings</span></span><br><span data-ttu-id="fd4a0-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="fd4a0-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-167">
        - Selection</span></span><br><span data-ttu-id="fd4a0-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-168">
        - Settings</span></span><br><span data-ttu-id="fd4a0-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-169">
        - TableBindings</span></span><br><span data-ttu-id="fd4a0-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-170">
        - TableCoercion</span></span><br><span data-ttu-id="fd4a0-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-171">
        - TextBindings</span></span><br><span data-ttu-id="fd4a0-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-173">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="fd4a0-173">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="fd4a0-174">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-174">- TaskPane</span></span><br><span data-ttu-id="fd4a0-175">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="fd4a0-175">
        - Content</span></span><br><span data-ttu-id="fd4a0-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="fd4a0-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fd4a0-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="fd4a0-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="fd4a0-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="fd4a0-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="fd4a0-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="fd4a0-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="fd4a0-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="fd4a0-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="fd4a0-186">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-186">- BindingEvents</span></span><br><span data-ttu-id="fd4a0-187">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-187">
        - CompressedFile</span></span><br><span data-ttu-id="fd4a0-188">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-188">
        - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-189">
        - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-189">
        - File</span></span><br><span data-ttu-id="fd4a0-190">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-190">
        - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-191">
        - MatrixBindings</span></span><br><span data-ttu-id="fd4a0-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="fd4a0-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-193">
        - Selection</span></span><br><span data-ttu-id="fd4a0-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-194">
        - Settings</span></span><br><span data-ttu-id="fd4a0-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-195">
        - TableBindings</span></span><br><span data-ttu-id="fd4a0-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-196">
        - TableCoercion</span></span><br><span data-ttu-id="fd4a0-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-197">
        - TextBindings</span></span><br><span data-ttu-id="fd4a0-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-199">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="fd4a0-199">Office for iPad</span></span></td>
    <td><span data-ttu-id="fd4a0-200">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-200">- TaskPane</span></span><br><span data-ttu-id="fd4a0-201">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="fd4a0-201">
        - Content</span></span></td>
    <td><span data-ttu-id="fd4a0-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fd4a0-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="fd4a0-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="fd4a0-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="fd4a0-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="fd4a0-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="fd4a0-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="fd4a0-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="fd4a0-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="fd4a0-211">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-211">- BindingEvents</span></span><br><span data-ttu-id="fd4a0-212">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-212">
        - CompressedFile</span></span><br><span data-ttu-id="fd4a0-213">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-213">
        - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-214">
        - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-214">
        - File</span></span><br><span data-ttu-id="fd4a0-215">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-215">
        - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-216">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-216">
        - MatrixBindings</span></span><br><span data-ttu-id="fd4a0-217">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-217">
        - MatrixCoercion</span></span><br><span data-ttu-id="fd4a0-218">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-218">
        - Selection</span></span><br><span data-ttu-id="fd4a0-219">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-219">
        - Settings</span></span><br><span data-ttu-id="fd4a0-220">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-220">
        - TableBindings</span></span><br><span data-ttu-id="fd4a0-221">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-221">
        - TableCoercion</span></span><br><span data-ttu-id="fd4a0-222">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-222">
        - TextBindings</span></span><br><span data-ttu-id="fd4a0-223">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-223">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-224">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="fd4a0-224">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="fd4a0-225">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-225">- TaskPane</span></span><br><span data-ttu-id="fd4a0-226">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="fd4a0-226">
        - Content</span></span></td>
    <td><span data-ttu-id="fd4a0-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fd4a0-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="fd4a0-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="fd4a0-229">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-229">- BindingEvents</span></span><br><span data-ttu-id="fd4a0-230">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-230">
        - CompressedFile</span></span><br><span data-ttu-id="fd4a0-231">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-231">
        - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-232">
        - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-232">
        - File</span></span><br><span data-ttu-id="fd4a0-233">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-233">
        - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-234">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-234">
        - MatrixBindings</span></span><br><span data-ttu-id="fd4a0-235">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-235">
        - MatrixCoercion</span></span><br><span data-ttu-id="fd4a0-236">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-236">
        - PdfFile</span></span><br><span data-ttu-id="fd4a0-237">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-237">
        - Selection</span></span><br><span data-ttu-id="fd4a0-238">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-238">
        - Settings</span></span><br><span data-ttu-id="fd4a0-239">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-239">
        - TableBindings</span></span><br><span data-ttu-id="fd4a0-240">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-240">
        - TableCoercion</span></span><br><span data-ttu-id="fd4a0-241">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-241">
        - TextBindings</span></span><br><span data-ttu-id="fd4a0-242">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-242">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-243">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="fd4a0-243">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="fd4a0-244">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-244">- TaskPane</span></span><br><span data-ttu-id="fd4a0-245">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="fd4a0-245">
        - Content</span></span><br><span data-ttu-id="fd4a0-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="fd4a0-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="fd4a0-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="fd4a0-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="fd4a0-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="fd4a0-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="fd4a0-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="fd4a0-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="fd4a0-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="fd4a0-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="fd4a0-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-256">- BindingEvents</span></span><br><span data-ttu-id="fd4a0-257">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-257">
        - CompressedFile</span></span><br><span data-ttu-id="fd4a0-258">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-258">
        - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-259">
        - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-259">
        - File</span></span><br><span data-ttu-id="fd4a0-260">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-260">
        - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-261">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-261">
        - MatrixBindings</span></span><br><span data-ttu-id="fd4a0-262">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-262">
        - MatrixCoercion</span></span><br><span data-ttu-id="fd4a0-263">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-263">
        - PdfFile</span></span><br><span data-ttu-id="fd4a0-264">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-264">
        - Selection</span></span><br><span data-ttu-id="fd4a0-265">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-265">
        - Settings</span></span><br><span data-ttu-id="fd4a0-266">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-266">
        - TableBindings</span></span><br><span data-ttu-id="fd4a0-267">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-267">
        - TableCoercion</span></span><br><span data-ttu-id="fd4a0-268">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-268">
        - TextBindings</span></span><br><span data-ttu-id="fd4a0-269">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-269">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="fd4a0-270">Outlook</span><span class="sxs-lookup"><span data-stu-id="fd4a0-270">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fd4a0-271">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="fd4a0-271">Platform</span></span></th>
    <th><span data-ttu-id="fd4a0-272">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="fd4a0-272">Extension points</span></span></th>
    <th><span data-ttu-id="fd4a0-273">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="fd4a0-273">API requirement sets</span></span></th>
    <th><span data-ttu-id="fd4a0-274"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-274"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-275">Office Online</span><span class="sxs-lookup"><span data-stu-id="fd4a0-275">Office Online</span></span></td>
    <td> <span data-ttu-id="fd4a0-276">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="fd4a0-276">- Mail Read</span></span><br><span data-ttu-id="fd4a0-277">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="fd4a0-277">
      - Mail Compose</span></span><br><span data-ttu-id="fd4a0-278">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-278">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-279">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-279">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fd4a0-280">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-280">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fd4a0-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fd4a0-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fd4a0-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="fd4a0-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="fd4a0-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="fd4a0-286">Non disponible</span><span class="sxs-lookup"><span data-stu-id="fd4a0-286">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-287">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="fd4a0-287">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="fd4a0-288">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="fd4a0-288">- Mail Read</span></span><br><span data-ttu-id="fd4a0-289">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="fd4a0-289">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="fd4a0-290">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-290">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fd4a0-291">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-291">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fd4a0-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fd4a0-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="fd4a0-294">Non disponible</span><span class="sxs-lookup"><span data-stu-id="fd4a0-294">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-295">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="fd4a0-295">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="fd4a0-296">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="fd4a0-296">- Mail Read</span></span><br><span data-ttu-id="fd4a0-297">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="fd4a0-297">
      - Mail Compose</span></span><br><span data-ttu-id="fd4a0-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="fd4a0-299">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="fd4a0-299">
      - Modules</span></span></td>
    <td> <span data-ttu-id="fd4a0-300">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-300">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fd4a0-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fd4a0-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fd4a0-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fd4a0-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="fd4a0-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="fd4a0-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="fd4a0-307">Non disponible</span><span class="sxs-lookup"><span data-stu-id="fd4a0-307">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-308">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="fd4a0-308">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="fd4a0-309">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="fd4a0-309">- Mail Read</span></span><br><span data-ttu-id="fd4a0-310">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="fd4a0-310">
      - Mail Compose</span></span><br><span data-ttu-id="fd4a0-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="fd4a0-312">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="fd4a0-312">
      - Modules</span></span></td>
    <td> <span data-ttu-id="fd4a0-313">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-313">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fd4a0-314">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-314">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fd4a0-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fd4a0-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fd4a0-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="fd4a0-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="fd4a0-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="fd4a0-320">Non disponible</span><span class="sxs-lookup"><span data-stu-id="fd4a0-320">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-321">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="fd4a0-321">Office for iOS</span></span></td>
    <td> <span data-ttu-id="fd4a0-322">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="fd4a0-322">- Mail Read</span></span><br><span data-ttu-id="fd4a0-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-324">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-324">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fd4a0-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fd4a0-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fd4a0-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fd4a0-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="fd4a0-329">Non disponible</span><span class="sxs-lookup"><span data-stu-id="fd4a0-329">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-330">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="fd4a0-330">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="fd4a0-331">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="fd4a0-331">- Mail Read</span></span><br><span data-ttu-id="fd4a0-332">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="fd4a0-332">
      - Mail Compose</span></span><br><span data-ttu-id="fd4a0-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-334">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-334">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fd4a0-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fd4a0-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fd4a0-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fd4a0-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="fd4a0-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="fd4a0-340">Non disponible</span><span class="sxs-lookup"><span data-stu-id="fd4a0-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-341">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="fd4a0-341">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="fd4a0-342">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="fd4a0-342">- Mail Read</span></span><br><span data-ttu-id="fd4a0-343">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="fd4a0-343">
      - Mail Compose</span></span><br><span data-ttu-id="fd4a0-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-345">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-345">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fd4a0-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fd4a0-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fd4a0-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fd4a0-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="fd4a0-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="fd4a0-351">Non disponible</span><span class="sxs-lookup"><span data-stu-id="fd4a0-351">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-352">Office pour Android</span><span class="sxs-lookup"><span data-stu-id="fd4a0-352">Office for Android</span></span></td>
    <td> <span data-ttu-id="fd4a0-353">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="fd4a0-353">- Mail Read</span></span><br><span data-ttu-id="fd4a0-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-355">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-355">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="fd4a0-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="fd4a0-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="fd4a0-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="fd4a0-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="fd4a0-360">Non disponible</span><span class="sxs-lookup"><span data-stu-id="fd4a0-360">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="fd4a0-361">Word</span><span class="sxs-lookup"><span data-stu-id="fd4a0-361">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fd4a0-362">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="fd4a0-362">Platform</span></span></th>
    <th><span data-ttu-id="fd4a0-363">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="fd4a0-363">Extension points</span></span></th>
    <th><span data-ttu-id="fd4a0-364">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="fd4a0-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="fd4a0-365"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-365"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-366">Office Online</span><span class="sxs-lookup"><span data-stu-id="fd4a0-366">Office Online</span></span></td>
    <td> <span data-ttu-id="fd4a0-367">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-367">- TaskPane</span></span><br><span data-ttu-id="fd4a0-368">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-368">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-369">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-369">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="fd4a0-370">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-370">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="fd4a0-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="fd4a0-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-373">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-373">- BindingEvents</span></span><br><span data-ttu-id="fd4a0-374">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fd4a0-374">
         - CustomXmlParts</span></span><br><span data-ttu-id="fd4a0-375">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-375">
         - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-376">
         - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-376">
         - File</span></span><br><span data-ttu-id="fd4a0-377">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-377">
         - HtmlCoercion</span></span><br><span data-ttu-id="fd4a0-378">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-378">
         - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-379">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-379">
         - MatrixBindings</span></span><br><span data-ttu-id="fd4a0-380">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-380">
         - MatrixCoercion</span></span><br><span data-ttu-id="fd4a0-381">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-381">
         - OoxmlCoercion</span></span><br><span data-ttu-id="fd4a0-382">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-382">
         - PdfFile</span></span><br><span data-ttu-id="fd4a0-383">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-383">
         - Selection</span></span><br><span data-ttu-id="fd4a0-384">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-384">
         - Settings</span></span><br><span data-ttu-id="fd4a0-385">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-385">
         - TableBindings</span></span><br><span data-ttu-id="fd4a0-386">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-386">
         - TableCoercion</span></span><br><span data-ttu-id="fd4a0-387">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-387">
         - TextBindings</span></span><br><span data-ttu-id="fd4a0-388">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-388">
         - TextCoercion</span></span><br><span data-ttu-id="fd4a0-389">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-389">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-390">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="fd4a0-390">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="fd4a0-391">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-391">- TaskPane</span></span></td>
    <td> <span data-ttu-id="fd4a0-392">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="fd4a0-392">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="fd4a0-393">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-393">- BindingEvents</span></span><br><span data-ttu-id="fd4a0-394">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-394">
         - CompressedFile</span></span><br><span data-ttu-id="fd4a0-395">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fd4a0-395">
         - CustomXmlParts</span></span><br><span data-ttu-id="fd4a0-396">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-396">
         - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-397">
         - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-397">
         - File</span></span><br><span data-ttu-id="fd4a0-398">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-398">
         - HtmlCoercion</span></span><br><span data-ttu-id="fd4a0-399">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-399">
         - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-400">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-400">
         - MatrixBindings</span></span><br><span data-ttu-id="fd4a0-401">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-401">
         - MatrixCoercion</span></span><br><span data-ttu-id="fd4a0-402">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-402">
         - OoxmlCoercion</span></span><br><span data-ttu-id="fd4a0-403">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-403">
         - PdfFile</span></span><br><span data-ttu-id="fd4a0-404">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-404">
         - Selection</span></span><br><span data-ttu-id="fd4a0-405">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-405">
         - Settings</span></span><br><span data-ttu-id="fd4a0-406">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-406">
         - TableBindings</span></span><br><span data-ttu-id="fd4a0-407">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-407">
         - TableCoercion</span></span><br><span data-ttu-id="fd4a0-408">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-408">
         - TextBindings</span></span><br><span data-ttu-id="fd4a0-409">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-409">
         - TextCoercion</span></span><br><span data-ttu-id="fd4a0-410">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-410">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-411">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="fd4a0-411">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="fd4a0-412">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-412">- TaskPane</span></span></td>
    <td> <span data-ttu-id="fd4a0-413">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-413">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="fd4a0-414">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="fd4a0-414">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="fd4a0-415">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-415">- BindingEvents</span></span><br><span data-ttu-id="fd4a0-416">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-416">
         - CompressedFile</span></span><br><span data-ttu-id="fd4a0-417">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fd4a0-417">
         - CustomXmlParts</span></span><br><span data-ttu-id="fd4a0-418">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-418">
         - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-419">
         - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-419">
         - File</span></span><br><span data-ttu-id="fd4a0-420">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-420">
         - HtmlCoercion</span></span><br><span data-ttu-id="fd4a0-421">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-421">
         - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-422">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-422">
         - MatrixBindings</span></span><br><span data-ttu-id="fd4a0-423">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-423">
         - MatrixCoercion</span></span><br><span data-ttu-id="fd4a0-424">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-424">
         - OoxmlCoercion</span></span><br><span data-ttu-id="fd4a0-425">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-425">
         - PdfFile</span></span><br><span data-ttu-id="fd4a0-426">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-426">
         - Selection</span></span><br><span data-ttu-id="fd4a0-427">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-427">
         - Settings</span></span><br><span data-ttu-id="fd4a0-428">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-428">
         - TableBindings</span></span><br><span data-ttu-id="fd4a0-429">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-429">
         - TableCoercion</span></span><br><span data-ttu-id="fd4a0-430">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-430">
         - TextBindings</span></span><br><span data-ttu-id="fd4a0-431">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-431">
         - TextCoercion</span></span><br><span data-ttu-id="fd4a0-432">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-432">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-433">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="fd4a0-433">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="fd4a0-434">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-434">- TaskPane</span></span><br><span data-ttu-id="fd4a0-435">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-435">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-436">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-436">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="fd4a0-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="fd4a0-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="fd4a0-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-440">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-440">- BindingEvents</span></span><br><span data-ttu-id="fd4a0-441">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-441">
         - CompressedFile</span></span><br><span data-ttu-id="fd4a0-442">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fd4a0-442">
         - CustomXmlParts</span></span><br><span data-ttu-id="fd4a0-443">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-443">
         - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-444">
         - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-444">
         - File</span></span><br><span data-ttu-id="fd4a0-445">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-445">
         - HtmlCoercion</span></span><br><span data-ttu-id="fd4a0-446">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-446">
         - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-447">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-447">
         - MatrixBindings</span></span><br><span data-ttu-id="fd4a0-448">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-448">
         - MatrixCoercion</span></span><br><span data-ttu-id="fd4a0-449">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-449">
         - OoxmlCoercion</span></span><br><span data-ttu-id="fd4a0-450">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-450">
         - PdfFile</span></span><br><span data-ttu-id="fd4a0-451">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-451">
         - Selection</span></span><br><span data-ttu-id="fd4a0-452">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-452">
         - Settings</span></span><br><span data-ttu-id="fd4a0-453">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-453">
         - TableBindings</span></span><br><span data-ttu-id="fd4a0-454">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-454">
         - TableCoercion</span></span><br><span data-ttu-id="fd4a0-455">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-455">
         - TextBindings</span></span><br><span data-ttu-id="fd4a0-456">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-456">
         - TextCoercion</span></span><br><span data-ttu-id="fd4a0-457">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-457">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-458">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="fd4a0-458">Office for iPad</span></span></td>
    <td> <span data-ttu-id="fd4a0-459">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-459">- TaskPane</span></span></td>
    <td> <span data-ttu-id="fd4a0-460">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-460">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="fd4a0-461">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-461">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="fd4a0-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="fd4a0-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="fd4a0-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="fd4a0-464">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-464">- BindingEvents</span></span><br><span data-ttu-id="fd4a0-465">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-465">
         - CompressedFile</span></span><br><span data-ttu-id="fd4a0-466">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fd4a0-466">
         - CustomXmlParts</span></span><br><span data-ttu-id="fd4a0-467">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-467">
         - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-468">
         - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-468">
         - File</span></span><br><span data-ttu-id="fd4a0-469">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-469">
         - HtmlCoercion</span></span><br><span data-ttu-id="fd4a0-470">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-470">
         - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-471">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-471">
         - MatrixBindings</span></span><br><span data-ttu-id="fd4a0-472">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-472">
         - MatrixCoercion</span></span><br><span data-ttu-id="fd4a0-473">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-473">
         - OoxmlCoercion</span></span><br><span data-ttu-id="fd4a0-474">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-474">
         - PdfFile</span></span><br><span data-ttu-id="fd4a0-475">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-475">
         - Selection</span></span><br><span data-ttu-id="fd4a0-476">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-476">
         - Settings</span></span><br><span data-ttu-id="fd4a0-477">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-477">
         - TableBindings</span></span><br><span data-ttu-id="fd4a0-478">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-478">
         - TableCoercion</span></span><br><span data-ttu-id="fd4a0-479">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-479">
         - TextBindings</span></span><br><span data-ttu-id="fd4a0-480">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-480">
         - TextCoercion</span></span><br><span data-ttu-id="fd4a0-481">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-481">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-482">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="fd4a0-482">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="fd4a0-483">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-483">- TaskPane</span></span></td>
    <td> <span data-ttu-id="fd4a0-484">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-484">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="fd4a0-485">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="fd4a0-485">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="fd4a0-486">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-486">- BindingEvents</span></span><br><span data-ttu-id="fd4a0-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-487">
         - CompressedFile</span></span><br><span data-ttu-id="fd4a0-488">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fd4a0-488">
         - CustomXmlParts</span></span><br><span data-ttu-id="fd4a0-489">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-489">
         - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-490">
         - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-490">
         - File</span></span><br><span data-ttu-id="fd4a0-491">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-491">
         - HtmlCoercion</span></span><br><span data-ttu-id="fd4a0-492">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-492">
         - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-493">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-493">
         - MatrixBindings</span></span><br><span data-ttu-id="fd4a0-494">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-494">
         - MatrixCoercion</span></span><br><span data-ttu-id="fd4a0-495">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-495">
         - OoxmlCoercion</span></span><br><span data-ttu-id="fd4a0-496">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-496">
         - PdfFile</span></span><br><span data-ttu-id="fd4a0-497">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-497">
         - Selection</span></span><br><span data-ttu-id="fd4a0-498">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-498">
         - Settings</span></span><br><span data-ttu-id="fd4a0-499">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-499">
         - TableBindings</span></span><br><span data-ttu-id="fd4a0-500">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-500">
         - TableCoercion</span></span><br><span data-ttu-id="fd4a0-501">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-501">
         - TextBindings</span></span><br><span data-ttu-id="fd4a0-502">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-502">
         - TextCoercion</span></span><br><span data-ttu-id="fd4a0-503">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-503">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-504">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="fd4a0-504">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="fd4a0-505">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-505">- TaskPane</span></span><br><span data-ttu-id="fd4a0-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="fd4a0-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="fd4a0-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="fd4a0-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="fd4a0-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="fd4a0-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-511">- BindingEvents</span></span><br><span data-ttu-id="fd4a0-512">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-512">
         - CompressedFile</span></span><br><span data-ttu-id="fd4a0-513">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="fd4a0-513">
         - CustomXmlParts</span></span><br><span data-ttu-id="fd4a0-514">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-514">
         - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-515">
         - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-515">
         - File</span></span><br><span data-ttu-id="fd4a0-516">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-516">
         - HtmlCoercion</span></span><br><span data-ttu-id="fd4a0-517">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-517">
         - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-518">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-518">
         - MatrixBindings</span></span><br><span data-ttu-id="fd4a0-519">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-519">
         - MatrixCoercion</span></span><br><span data-ttu-id="fd4a0-520">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-520">
         - OoxmlCoercion</span></span><br><span data-ttu-id="fd4a0-521">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-521">
         - PdfFile</span></span><br><span data-ttu-id="fd4a0-522">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-522">
         - Selection</span></span><br><span data-ttu-id="fd4a0-523">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-523">
         - Settings</span></span><br><span data-ttu-id="fd4a0-524">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-524">
         - TableBindings</span></span><br><span data-ttu-id="fd4a0-525">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-525">
         - TableCoercion</span></span><br><span data-ttu-id="fd4a0-526">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-526">
         - TextBindings</span></span><br><span data-ttu-id="fd4a0-527">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-527">
         - TextCoercion</span></span><br><span data-ttu-id="fd4a0-528">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-528">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="fd4a0-529">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="fd4a0-529">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fd4a0-530">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="fd4a0-530">Platform</span></span></th>
    <th><span data-ttu-id="fd4a0-531">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="fd4a0-531">Extension points</span></span></th>
    <th><span data-ttu-id="fd4a0-532">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="fd4a0-532">API requirement sets</span></span></th>
    <th><span data-ttu-id="fd4a0-533"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-533"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-534">Office Online</span><span class="sxs-lookup"><span data-stu-id="fd4a0-534">Office Online</span></span></td>
    <td> <span data-ttu-id="fd4a0-535">- Contenu</span><span class="sxs-lookup"><span data-stu-id="fd4a0-535">- Content</span></span><br><span data-ttu-id="fd4a0-536">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-536">
         - TaskPane</span></span><br><span data-ttu-id="fd4a0-537">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-537">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-538">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-538">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-539">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="fd4a0-539">- ActiveView</span></span><br><span data-ttu-id="fd4a0-540">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-540">
         - CompressedFile</span></span><br><span data-ttu-id="fd4a0-541">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-541">
         - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-542">
         - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-542">
         - File</span></span><br><span data-ttu-id="fd4a0-543">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-543">
         - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-544">
         - PdfFile</span></span><br><span data-ttu-id="fd4a0-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-545">
         - Selection</span></span><br><span data-ttu-id="fd4a0-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-546">
         - Settings</span></span><br><span data-ttu-id="fd4a0-547">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-547">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-548">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="fd4a0-548">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="fd4a0-549">- Contenu</span><span class="sxs-lookup"><span data-stu-id="fd4a0-549">- Content</span></span><br><span data-ttu-id="fd4a0-550">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-550">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="fd4a0-551">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="fd4a0-551">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="fd4a0-552">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="fd4a0-552">- ActiveView</span></span><br><span data-ttu-id="fd4a0-553">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-553">
         - CompressedFile</span></span><br><span data-ttu-id="fd4a0-554">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-554">
         - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-555">
         - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-555">
         - File</span></span><br><span data-ttu-id="fd4a0-556">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-556">
         - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-557">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-557">
         - PdfFile</span></span><br><span data-ttu-id="fd4a0-558">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-558">
         - Selection</span></span><br><span data-ttu-id="fd4a0-559">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-559">
         - Settings</span></span><br><span data-ttu-id="fd4a0-560">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-560">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-561">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="fd4a0-561">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="fd4a0-562">- Contenu</span><span class="sxs-lookup"><span data-stu-id="fd4a0-562">- Content</span></span><br><span data-ttu-id="fd4a0-563">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-563">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="fd4a0-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="fd4a0-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="fd4a0-565">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="fd4a0-565">- ActiveView</span></span><br><span data-ttu-id="fd4a0-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-566">
         - CompressedFile</span></span><br><span data-ttu-id="fd4a0-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-567">
         - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-568">
         - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-568">
         - File</span></span><br><span data-ttu-id="fd4a0-569">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-569">
         - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-570">
         - PdfFile</span></span><br><span data-ttu-id="fd4a0-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-571">
         - Selection</span></span><br><span data-ttu-id="fd4a0-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-572">
         - Settings</span></span><br><span data-ttu-id="fd4a0-573">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-573">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-574">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="fd4a0-574">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="fd4a0-575">- Contenu</span><span class="sxs-lookup"><span data-stu-id="fd4a0-575">- Content</span></span><br><span data-ttu-id="fd4a0-576">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-576">
         - TaskPane</span></span><br><span data-ttu-id="fd4a0-577">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-577">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-578">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-578">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-579">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="fd4a0-579">- ActiveView</span></span><br><span data-ttu-id="fd4a0-580">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-580">
         - CompressedFile</span></span><br><span data-ttu-id="fd4a0-581">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-581">
         - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-582">
         - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-582">
         - File</span></span><br><span data-ttu-id="fd4a0-583">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-583">
         - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-584">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-584">
         - PdfFile</span></span><br><span data-ttu-id="fd4a0-585">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-585">
         - Selection</span></span><br><span data-ttu-id="fd4a0-586">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-586">
         - Settings</span></span><br><span data-ttu-id="fd4a0-587">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-587">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-588">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="fd4a0-588">Office for iPad</span></span></td>
    <td> <span data-ttu-id="fd4a0-589">- Contenu</span><span class="sxs-lookup"><span data-stu-id="fd4a0-589">- Content</span></span><br><span data-ttu-id="fd4a0-590">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-590">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="fd4a0-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="fd4a0-592">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="fd4a0-592">- ActiveView</span></span><br><span data-ttu-id="fd4a0-593">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-593">
         - CompressedFile</span></span><br><span data-ttu-id="fd4a0-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-594">
         - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-595">
         - File</span></span><br><span data-ttu-id="fd4a0-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-596">
         - PdfFile</span></span><br><span data-ttu-id="fd4a0-597">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-597">
         - Selection</span></span><br><span data-ttu-id="fd4a0-598">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-598">
         - Settings</span></span><br><span data-ttu-id="fd4a0-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-599">
         - TextCoercion</span></span><br><span data-ttu-id="fd4a0-600">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-600">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-601">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="fd4a0-601">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="fd4a0-602">- Contenu</span><span class="sxs-lookup"><span data-stu-id="fd4a0-602">- Content</span></span><br><span data-ttu-id="fd4a0-603">Volet Office 
         -/td></span><span class="sxs-lookup"><span data-stu-id="fd4a0-603">
         - TaskPane/td></span></span> <td> <span data-ttu-id="fd4a0-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="fd4a0-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="fd4a0-605">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="fd4a0-605">- ActiveView</span></span><br><span data-ttu-id="fd4a0-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-606">
         - CompressedFile</span></span><br><span data-ttu-id="fd4a0-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-607">
         - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-608">
         - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-608">
         - File</span></span><br><span data-ttu-id="fd4a0-609">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-609">
         - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-610">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-610">
         - PdfFile</span></span><br><span data-ttu-id="fd4a0-611">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-611">
         - Selection</span></span><br><span data-ttu-id="fd4a0-612">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-612">
         - Settings</span></span><br><span data-ttu-id="fd4a0-613">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-613">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-614">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="fd4a0-614">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="fd4a0-615">- Contenu</span><span class="sxs-lookup"><span data-stu-id="fd4a0-615">- Content</span></span><br><span data-ttu-id="fd4a0-616">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-616">
         - TaskPane</span></span><br><span data-ttu-id="fd4a0-617">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-617">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-619">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="fd4a0-619">- ActiveView</span></span><br><span data-ttu-id="fd4a0-620">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-620">
         - CompressedFile</span></span><br><span data-ttu-id="fd4a0-621">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-621">
         - DocumentEvents</span></span><br><span data-ttu-id="fd4a0-622">
         - File</span><span class="sxs-lookup"><span data-stu-id="fd4a0-622">
         - File</span></span><br><span data-ttu-id="fd4a0-623">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-623">
         - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-624">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="fd4a0-624">
         - PdfFile</span></span><br><span data-ttu-id="fd4a0-625">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-625">
         - Selection</span></span><br><span data-ttu-id="fd4a0-626">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-626">
         - Settings</span></span><br><span data-ttu-id="fd4a0-627">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-627">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="fd4a0-628">OneNote</span><span class="sxs-lookup"><span data-stu-id="fd4a0-628">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fd4a0-629">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="fd4a0-629">Platform</span></span></th>
    <th><span data-ttu-id="fd4a0-630">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="fd4a0-630">Extension points</span></span></th>
    <th><span data-ttu-id="fd4a0-631">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="fd4a0-631">API requirement sets</span></span></th>
    <th><span data-ttu-id="fd4a0-632"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-632"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-633">Office Online</span><span class="sxs-lookup"><span data-stu-id="fd4a0-633">Office Online</span></span></td>
    <td> <span data-ttu-id="fd4a0-634">- Contenu</span><span class="sxs-lookup"><span data-stu-id="fd4a0-634">- Content</span></span><br><span data-ttu-id="fd4a0-635">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-635">
         - TaskPane</span></span><br><span data-ttu-id="fd4a0-636">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-636">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-637">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-637">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="fd4a0-638">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-638">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-639">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="fd4a0-639">- DocumentEvents</span></span><br><span data-ttu-id="fd4a0-640">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-640">
         - HtmlCoercion</span></span><br><span data-ttu-id="fd4a0-641">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-641">
         - ImageCoercion</span></span><br><span data-ttu-id="fd4a0-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="fd4a0-642">
         - Settings</span></span><br><span data-ttu-id="fd4a0-643">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-643">
         - TextCoercion</span></span></td>
  </tr>
</table><span data-ttu-id="fd4a0-644">
\*&ast; : ajouté avec les mises à jour après la publication.*

</span><span class="sxs-lookup"><span data-stu-id="fd4a0-644">
\*&ast; - Added with post-release updates.*

</span></span><br/>

## <a name="project"></a><span data-ttu-id="fd4a0-645">Project</span><span class="sxs-lookup"><span data-stu-id="fd4a0-645">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="fd4a0-646">Plateforme</span><span class="sxs-lookup"><span data-stu-id="fd4a0-646">Platform</span></span></th>
    <th><span data-ttu-id="fd4a0-647">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="fd4a0-647">Extension points</span></span></th>
    <th><span data-ttu-id="fd4a0-648">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="fd4a0-648">API requirement sets</span></span></th>
    <th><span data-ttu-id="fd4a0-649"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-649"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-650">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="fd4a0-650">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="fd4a0-651">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-651">- TaskPane</span></span></td>
    <td> <span data-ttu-id="fd4a0-652">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-652">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-653">- Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-653">- Selection</span></span><br><span data-ttu-id="fd4a0-654">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-654">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-655">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="fd4a0-655">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="fd4a0-656">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-656">- TaskPane</span></span></td>
    <td> <span data-ttu-id="fd4a0-657">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-657">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-658">- Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-658">- Selection</span></span><br><span data-ttu-id="fd4a0-659">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-659">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="fd4a0-660">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="fd4a0-660">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="fd4a0-661">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-661">- TaskPane</span></span></td>
    <td> <span data-ttu-id="fd4a0-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="fd4a0-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="fd4a0-663">- Selection</span><span class="sxs-lookup"><span data-stu-id="fd4a0-663">- Selection</span></span><br><span data-ttu-id="fd4a0-664">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="fd4a0-664">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="fd4a0-665">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="fd4a0-665">See also</span></span>

- [<span data-ttu-id="fd4a0-666">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-666">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="fd4a0-667">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="fd4a0-667">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="fd4a0-668">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="fd4a0-668">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="fd4a0-669">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="fd4a0-669">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
