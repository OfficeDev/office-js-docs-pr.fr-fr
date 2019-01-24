---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint, OneNote et Project.
ms.date: 11/07/2018
localization_priority: Priority
ms.openlocfilehash: 9f8b94483d22f24dcb0a6a2ad99df6167533133f
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388338"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="8da31-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="8da31-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="8da31-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="8da31-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="8da31-p102">Le numéro de build pour Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que les ensembles de conditions requises des API communes, ExcelApi 1.1 et WordApi 1.1.</span><span class="sxs-lookup"><span data-stu-id="8da31-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="8da31-108">Excel</span><span class="sxs-lookup"><span data-stu-id="8da31-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="8da31-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="8da31-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="8da31-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="8da31-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="8da31-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="8da31-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="8da31-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="8da31-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="8da31-113">Office Online</span></span></td>
    <td> <span data-ttu-id="8da31-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-114">- TaskPane</span></span><br><span data-ttu-id="8da31-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="8da31-115">
        - Content</span></span><br><span data-ttu-id="8da31-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="8da31-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8da31-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8da31-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8da31-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8da31-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8da31-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8da31-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8da31-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8da31-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8da31-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8da31-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8da31-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8da31-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8da31-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8da31-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8da31-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-126">
        - BindingEvents</span></span><br><span data-ttu-id="8da31-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-127">
        - CompressedFile</span></span><br><span data-ttu-id="8da31-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-128">
        - DocumentEvents</span></span><br><span data-ttu-id="8da31-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="8da31-129">
        - File</span></span><br><span data-ttu-id="8da31-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-130">
        - MatrixBindings</span></span><br><span data-ttu-id="8da31-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="8da31-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-132">
        - Selection</span></span><br><span data-ttu-id="8da31-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-133">
        - Settings</span></span><br><span data-ttu-id="8da31-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-134">
        - TableBindings</span></span><br><span data-ttu-id="8da31-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-135">
        - TableCoercion</span></span><br><span data-ttu-id="8da31-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-136">
        - TextBindings</span></span><br><span data-ttu-id="8da31-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-138">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="8da31-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="8da31-139">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-139">
        - TaskPane</span></span><br><span data-ttu-id="8da31-140">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="8da31-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="8da31-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8da31-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-142">
        - BindingEvents</span></span><br><span data-ttu-id="8da31-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-143">
        - CompressedFile</span></span><br><span data-ttu-id="8da31-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-144">
        - DocumentEvents</span></span><br><span data-ttu-id="8da31-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="8da31-145">
        - File</span></span><br><span data-ttu-id="8da31-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-146">
        - ImageCoercion</span></span><br><span data-ttu-id="8da31-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-147">
        - MatrixBindings</span></span><br><span data-ttu-id="8da31-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="8da31-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-149">
        - Selection</span></span><br><span data-ttu-id="8da31-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-150">
        - Settings</span></span><br><span data-ttu-id="8da31-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-151">
        - TableBindings</span></span><br><span data-ttu-id="8da31-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-152">
        - TableCoercion</span></span><br><span data-ttu-id="8da31-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-153">
        - TextBindings</span></span><br><span data-ttu-id="8da31-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-155">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="8da31-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="8da31-156">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-156">- TaskPane</span></span><br><span data-ttu-id="8da31-157">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="8da31-157">
        - Content</span></span><br><span data-ttu-id="8da31-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="8da31-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8da31-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8da31-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8da31-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8da31-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8da31-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8da31-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8da31-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8da31-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8da31-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8da31-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8da31-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8da31-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8da31-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8da31-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-168">- BindingEvents</span></span><br><span data-ttu-id="8da31-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-169">
        - CompressedFile</span></span><br><span data-ttu-id="8da31-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-170">
        - DocumentEvents</span></span><br><span data-ttu-id="8da31-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="8da31-171">
        - File</span></span><br><span data-ttu-id="8da31-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-172">
        - ImageCoercion</span></span><br><span data-ttu-id="8da31-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-173">
        - MatrixBindings</span></span><br><span data-ttu-id="8da31-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-174">
        - MatrixCoercion</span></span><br><span data-ttu-id="8da31-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-175">
        - Selection</span></span><br><span data-ttu-id="8da31-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-176">
        - Settings</span></span><br><span data-ttu-id="8da31-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-177">
        - TableBindings</span></span><br><span data-ttu-id="8da31-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-178">
        - TableCoercion</span></span><br><span data-ttu-id="8da31-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-179">
        - TextBindings</span></span><br><span data-ttu-id="8da31-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-180">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-181">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="8da31-181">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="8da31-182">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-182">- TaskPane</span></span><br><span data-ttu-id="8da31-183">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="8da31-183">
        - Content</span></span><br><span data-ttu-id="8da31-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="8da31-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8da31-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8da31-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8da31-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8da31-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8da31-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8da31-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8da31-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8da31-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8da31-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8da31-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8da31-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8da31-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8da31-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8da31-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-194">- BindingEvents</span></span><br><span data-ttu-id="8da31-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-195">
        - CompressedFile</span></span><br><span data-ttu-id="8da31-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-196">
        - DocumentEvents</span></span><br><span data-ttu-id="8da31-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="8da31-197">
        - File</span></span><br><span data-ttu-id="8da31-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-198">
        - ImageCoercion</span></span><br><span data-ttu-id="8da31-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-199">
        - MatrixBindings</span></span><br><span data-ttu-id="8da31-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="8da31-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-201">
        - Selection</span></span><br><span data-ttu-id="8da31-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-202">
        - Settings</span></span><br><span data-ttu-id="8da31-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-203">
        - TableBindings</span></span><br><span data-ttu-id="8da31-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-204">
        - TableCoercion</span></span><br><span data-ttu-id="8da31-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-205">
        - TextBindings</span></span><br><span data-ttu-id="8da31-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-207">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="8da31-207">Office for iPad</span></span></td>
    <td><span data-ttu-id="8da31-208">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-208">- TaskPane</span></span><br><span data-ttu-id="8da31-209">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="8da31-209">
        - Content</span></span></td>
    <td><span data-ttu-id="8da31-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8da31-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8da31-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8da31-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8da31-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8da31-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8da31-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8da31-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8da31-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8da31-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8da31-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8da31-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8da31-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8da31-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8da31-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-219">- BindingEvents</span></span><br><span data-ttu-id="8da31-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-220">
        - CompressedFile</span></span><br><span data-ttu-id="8da31-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-221">
        - DocumentEvents</span></span><br><span data-ttu-id="8da31-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="8da31-222">
        - File</span></span><br><span data-ttu-id="8da31-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-223">
        - ImageCoercion</span></span><br><span data-ttu-id="8da31-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-224">
        - MatrixBindings</span></span><br><span data-ttu-id="8da31-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="8da31-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-226">
        - Selection</span></span><br><span data-ttu-id="8da31-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-227">
        - Settings</span></span><br><span data-ttu-id="8da31-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-228">
        - TableBindings</span></span><br><span data-ttu-id="8da31-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-229">
        - TableCoercion</span></span><br><span data-ttu-id="8da31-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-230">
        - TextBindings</span></span><br><span data-ttu-id="8da31-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-232">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="8da31-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="8da31-233">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-233">- TaskPane</span></span><br><span data-ttu-id="8da31-234">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="8da31-234">
        - Content</span></span><br><span data-ttu-id="8da31-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="8da31-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8da31-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8da31-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8da31-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8da31-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8da31-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8da31-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8da31-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8da31-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8da31-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8da31-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8da31-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8da31-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8da31-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8da31-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-245">- BindingEvents</span></span><br><span data-ttu-id="8da31-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-246">
        - CompressedFile</span></span><br><span data-ttu-id="8da31-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-247">
        - DocumentEvents</span></span><br><span data-ttu-id="8da31-248">
        - File</span><span class="sxs-lookup"><span data-stu-id="8da31-248">
        - File</span></span><br><span data-ttu-id="8da31-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-249">
        - ImageCoercion</span></span><br><span data-ttu-id="8da31-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-250">
        - MatrixBindings</span></span><br><span data-ttu-id="8da31-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-251">
        - MatrixCoercion</span></span><br><span data-ttu-id="8da31-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8da31-252">
        - PdfFile</span></span><br><span data-ttu-id="8da31-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-253">
        - Selection</span></span><br><span data-ttu-id="8da31-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-254">
        - Settings</span></span><br><span data-ttu-id="8da31-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-255">
        - TableBindings</span></span><br><span data-ttu-id="8da31-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-256">
        - TableCoercion</span></span><br><span data-ttu-id="8da31-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-257">
        - TextBindings</span></span><br><span data-ttu-id="8da31-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-259">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="8da31-259">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="8da31-260">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-260">- TaskPane</span></span><br><span data-ttu-id="8da31-261">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="8da31-261">
        - Content</span></span><br><span data-ttu-id="8da31-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="8da31-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8da31-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8da31-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8da31-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8da31-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8da31-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8da31-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8da31-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8da31-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8da31-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8da31-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8da31-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8da31-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8da31-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8da31-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-272">- BindingEvents</span></span><br><span data-ttu-id="8da31-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-273">
        - CompressedFile</span></span><br><span data-ttu-id="8da31-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-274">
        - DocumentEvents</span></span><br><span data-ttu-id="8da31-275">
        - File</span><span class="sxs-lookup"><span data-stu-id="8da31-275">
        - File</span></span><br><span data-ttu-id="8da31-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-276">
        - ImageCoercion</span></span><br><span data-ttu-id="8da31-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-277">
        - MatrixBindings</span></span><br><span data-ttu-id="8da31-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-278">
        - MatrixCoercion</span></span><br><span data-ttu-id="8da31-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8da31-279">
        - PdfFile</span></span><br><span data-ttu-id="8da31-280">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-280">
        - Selection</span></span><br><span data-ttu-id="8da31-281">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-281">
        - Settings</span></span><br><span data-ttu-id="8da31-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-282">
        - TableBindings</span></span><br><span data-ttu-id="8da31-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-283">
        - TableCoercion</span></span><br><span data-ttu-id="8da31-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-284">
        - TextBindings</span></span><br><span data-ttu-id="8da31-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-285">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="8da31-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="8da31-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8da31-287">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="8da31-287">Platform</span></span></th>
    <th><span data-ttu-id="8da31-288">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="8da31-288">Extension points</span></span></th>
    <th><span data-ttu-id="8da31-289">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="8da31-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="8da31-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="8da31-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="8da31-291">Office Online</span></span></td>
    <td> <span data-ttu-id="8da31-292">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="8da31-292">- Mail Read</span></span><br><span data-ttu-id="8da31-293">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="8da31-293">
      - Mail Compose</span></span><br><span data-ttu-id="8da31-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8da31-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8da31-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8da31-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8da31-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8da31-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8da31-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8da31-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8da31-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8da31-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="8da31-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8da31-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="8da31-302">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8da31-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-303">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="8da31-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="8da31-304">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="8da31-304">- Mail Read</span></span><br><span data-ttu-id="8da31-305">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="8da31-305">
      - Mail Compose</span></span><br><span data-ttu-id="8da31-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8da31-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8da31-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8da31-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8da31-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8da31-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="8da31-311">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8da31-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-312">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="8da31-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="8da31-313">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="8da31-313">- Mail Read</span></span><br><span data-ttu-id="8da31-314">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="8da31-314">
      - Mail Compose</span></span><br><span data-ttu-id="8da31-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="8da31-316">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="8da31-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="8da31-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8da31-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8da31-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8da31-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8da31-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8da31-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8da31-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8da31-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8da31-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="8da31-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8da31-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="8da31-324">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8da31-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-325">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="8da31-325">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="8da31-326">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="8da31-326">- Mail Read</span></span><br><span data-ttu-id="8da31-327">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="8da31-327">
      - Mail Compose</span></span><br><span data-ttu-id="8da31-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="8da31-329">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="8da31-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="8da31-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8da31-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8da31-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8da31-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8da31-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8da31-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8da31-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8da31-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8da31-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="8da31-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8da31-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="8da31-337">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8da31-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-338">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="8da31-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="8da31-339">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="8da31-339">- Mail Read</span></span><br><span data-ttu-id="8da31-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8da31-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8da31-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8da31-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8da31-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8da31-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8da31-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8da31-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="8da31-346">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8da31-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-347">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="8da31-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="8da31-348">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="8da31-348">- Mail Read</span></span><br><span data-ttu-id="8da31-349">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="8da31-349">
      - Mail Compose</span></span><br><span data-ttu-id="8da31-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8da31-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8da31-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8da31-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8da31-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8da31-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8da31-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8da31-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8da31-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8da31-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="8da31-357">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8da31-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-358">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="8da31-358">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="8da31-359">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="8da31-359">- Mail Read</span></span><br><span data-ttu-id="8da31-360">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="8da31-360">
      - Mail Compose</span></span><br><span data-ttu-id="8da31-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8da31-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8da31-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8da31-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8da31-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8da31-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8da31-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8da31-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8da31-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8da31-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="8da31-368">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8da31-368">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-369">Office pour Android</span><span class="sxs-lookup"><span data-stu-id="8da31-369">Office for Android</span></span></td>
    <td> <span data-ttu-id="8da31-370">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="8da31-370">- Mail Read</span></span><br><span data-ttu-id="8da31-371">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-371">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8da31-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8da31-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8da31-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8da31-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8da31-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8da31-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8da31-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="8da31-377">Non disponible</span><span class="sxs-lookup"><span data-stu-id="8da31-377">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="8da31-378">Word</span><span class="sxs-lookup"><span data-stu-id="8da31-378">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8da31-379">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="8da31-379">Platform</span></span></th>
    <th><span data-ttu-id="8da31-380">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="8da31-380">Extension points</span></span></th>
    <th><span data-ttu-id="8da31-381">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="8da31-381">API requirement sets</span></span></th>
    <th><span data-ttu-id="8da31-382"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="8da31-382"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-383">Office Online</span><span class="sxs-lookup"><span data-stu-id="8da31-383">Office Online</span></span></td>
    <td> <span data-ttu-id="8da31-384">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-384">- TaskPane</span></span><br><span data-ttu-id="8da31-385">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-385">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8da31-386">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-386">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="8da31-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="8da31-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="8da31-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8da31-390">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-390">- BindingEvents</span></span><br><span data-ttu-id="8da31-391">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8da31-391">
         - CustomXmlParts</span></span><br><span data-ttu-id="8da31-392">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-392">
         - DocumentEvents</span></span><br><span data-ttu-id="8da31-393">
         - File</span><span class="sxs-lookup"><span data-stu-id="8da31-393">
         - File</span></span><br><span data-ttu-id="8da31-394">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-394">
         - HtmlCoercion</span></span><br><span data-ttu-id="8da31-395">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-395">
         - ImageCoercion</span></span><br><span data-ttu-id="8da31-396">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-396">
         - MatrixBindings</span></span><br><span data-ttu-id="8da31-397">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-397">
         - MatrixCoercion</span></span><br><span data-ttu-id="8da31-398">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-398">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8da31-399">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8da31-399">
         - PdfFile</span></span><br><span data-ttu-id="8da31-400">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-400">
         - Selection</span></span><br><span data-ttu-id="8da31-401">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-401">
         - Settings</span></span><br><span data-ttu-id="8da31-402">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-402">
         - TableBindings</span></span><br><span data-ttu-id="8da31-403">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-403">
         - TableCoercion</span></span><br><span data-ttu-id="8da31-404">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-404">
         - TextBindings</span></span><br><span data-ttu-id="8da31-405">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-405">
         - TextCoercion</span></span><br><span data-ttu-id="8da31-406">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8da31-406">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-407">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="8da31-407">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="8da31-408">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-408">- TaskPane</span></span></td>
    <td> <span data-ttu-id="8da31-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8da31-410">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-410">- BindingEvents</span></span><br><span data-ttu-id="8da31-411">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-411">
         - CompressedFile</span></span><br><span data-ttu-id="8da31-412">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8da31-412">
         - CustomXmlParts</span></span><br><span data-ttu-id="8da31-413">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-413">
         - DocumentEvents</span></span><br><span data-ttu-id="8da31-414">
         - File</span><span class="sxs-lookup"><span data-stu-id="8da31-414">
         - File</span></span><br><span data-ttu-id="8da31-415">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-415">
         - HtmlCoercion</span></span><br><span data-ttu-id="8da31-416">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-416">
         - ImageCoercion</span></span><br><span data-ttu-id="8da31-417">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-417">
         - MatrixBindings</span></span><br><span data-ttu-id="8da31-418">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-418">
         - MatrixCoercion</span></span><br><span data-ttu-id="8da31-419">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-419">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8da31-420">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8da31-420">
         - PdfFile</span></span><br><span data-ttu-id="8da31-421">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-421">
         - Selection</span></span><br><span data-ttu-id="8da31-422">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-422">
         - Settings</span></span><br><span data-ttu-id="8da31-423">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-423">
         - TableBindings</span></span><br><span data-ttu-id="8da31-424">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-424">
         - TableCoercion</span></span><br><span data-ttu-id="8da31-425">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-425">
         - TextBindings</span></span><br><span data-ttu-id="8da31-426">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-426">
         - TextCoercion</span></span><br><span data-ttu-id="8da31-427">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8da31-427">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-428">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="8da31-428">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="8da31-429">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-429">- TaskPane</span></span><br><span data-ttu-id="8da31-430">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-430">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8da31-431">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-431">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="8da31-432">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-432">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="8da31-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="8da31-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8da31-435">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-435">- BindingEvents</span></span><br><span data-ttu-id="8da31-436">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-436">
         - CompressedFile</span></span><br><span data-ttu-id="8da31-437">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8da31-437">
         - CustomXmlParts</span></span><br><span data-ttu-id="8da31-438">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-438">
         - DocumentEvents</span></span><br><span data-ttu-id="8da31-439">
         - File</span><span class="sxs-lookup"><span data-stu-id="8da31-439">
         - File</span></span><br><span data-ttu-id="8da31-440">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-440">
         - HtmlCoercion</span></span><br><span data-ttu-id="8da31-441">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-441">
         - ImageCoercion</span></span><br><span data-ttu-id="8da31-442">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-442">
         - MatrixBindings</span></span><br><span data-ttu-id="8da31-443">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-443">
         - MatrixCoercion</span></span><br><span data-ttu-id="8da31-444">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-444">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8da31-445">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8da31-445">
         - PdfFile</span></span><br><span data-ttu-id="8da31-446">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-446">
         - Selection</span></span><br><span data-ttu-id="8da31-447">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-447">
         - Settings</span></span><br><span data-ttu-id="8da31-448">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-448">
         - TableBindings</span></span><br><span data-ttu-id="8da31-449">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-449">
         - TableCoercion</span></span><br><span data-ttu-id="8da31-450">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-450">
         - TextBindings</span></span><br><span data-ttu-id="8da31-451">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-451">
         - TextCoercion</span></span><br><span data-ttu-id="8da31-452">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8da31-452">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-453">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="8da31-453">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="8da31-454">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-454">- TaskPane</span></span><br><span data-ttu-id="8da31-455">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-455">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8da31-456">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-456">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="8da31-457">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-457">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="8da31-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="8da31-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8da31-460">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-460">- BindingEvents</span></span><br><span data-ttu-id="8da31-461">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-461">
         - CompressedFile</span></span><br><span data-ttu-id="8da31-462">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8da31-462">
         - CustomXmlParts</span></span><br><span data-ttu-id="8da31-463">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-463">
         - DocumentEvents</span></span><br><span data-ttu-id="8da31-464">
         - File</span><span class="sxs-lookup"><span data-stu-id="8da31-464">
         - File</span></span><br><span data-ttu-id="8da31-465">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-465">
         - HtmlCoercion</span></span><br><span data-ttu-id="8da31-466">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-466">
         - ImageCoercion</span></span><br><span data-ttu-id="8da31-467">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-467">
         - MatrixBindings</span></span><br><span data-ttu-id="8da31-468">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-468">
         - MatrixCoercion</span></span><br><span data-ttu-id="8da31-469">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-469">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8da31-470">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8da31-470">
         - PdfFile</span></span><br><span data-ttu-id="8da31-471">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-471">
         - Selection</span></span><br><span data-ttu-id="8da31-472">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-472">
         - Settings</span></span><br><span data-ttu-id="8da31-473">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-473">
         - TableBindings</span></span><br><span data-ttu-id="8da31-474">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-474">
         - TableCoercion</span></span><br><span data-ttu-id="8da31-475">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-475">
         - TextBindings</span></span><br><span data-ttu-id="8da31-476">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-476">
         - TextCoercion</span></span><br><span data-ttu-id="8da31-477">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8da31-477">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-478">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="8da31-478">Office for iPad</span></span></td>
    <td> <span data-ttu-id="8da31-479">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-479">- TaskPane</span></span></td>
    <td> <span data-ttu-id="8da31-480">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-480">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="8da31-481">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-481">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="8da31-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="8da31-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="8da31-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="8da31-484">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-484">- BindingEvents</span></span><br><span data-ttu-id="8da31-485">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-485">
         - CompressedFile</span></span><br><span data-ttu-id="8da31-486">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8da31-486">
         - CustomXmlParts</span></span><br><span data-ttu-id="8da31-487">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-487">
         - DocumentEvents</span></span><br><span data-ttu-id="8da31-488">
         - File</span><span class="sxs-lookup"><span data-stu-id="8da31-488">
         - File</span></span><br><span data-ttu-id="8da31-489">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-489">
         - HtmlCoercion</span></span><br><span data-ttu-id="8da31-490">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-490">
         - ImageCoercion</span></span><br><span data-ttu-id="8da31-491">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-491">
         - MatrixBindings</span></span><br><span data-ttu-id="8da31-492">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-492">
         - MatrixCoercion</span></span><br><span data-ttu-id="8da31-493">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-493">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8da31-494">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8da31-494">
         - PdfFile</span></span><br><span data-ttu-id="8da31-495">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-495">
         - Selection</span></span><br><span data-ttu-id="8da31-496">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-496">
         - Settings</span></span><br><span data-ttu-id="8da31-497">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-497">
         - TableBindings</span></span><br><span data-ttu-id="8da31-498">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-498">
         - TableCoercion</span></span><br><span data-ttu-id="8da31-499">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-499">
         - TextBindings</span></span><br><span data-ttu-id="8da31-500">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-500">
         - TextCoercion</span></span><br><span data-ttu-id="8da31-501">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8da31-501">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-502">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="8da31-502">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="8da31-503">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-503">- TaskPane</span></span><br><span data-ttu-id="8da31-504">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-504">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8da31-505">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-505">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="8da31-506">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-506">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="8da31-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="8da31-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="8da31-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="8da31-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-509">- BindingEvents</span></span><br><span data-ttu-id="8da31-510">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-510">
         - CompressedFile</span></span><br><span data-ttu-id="8da31-511">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8da31-511">
         - CustomXmlParts</span></span><br><span data-ttu-id="8da31-512">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-512">
         - DocumentEvents</span></span><br><span data-ttu-id="8da31-513">
         - File</span><span class="sxs-lookup"><span data-stu-id="8da31-513">
         - File</span></span><br><span data-ttu-id="8da31-514">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-514">
         - HtmlCoercion</span></span><br><span data-ttu-id="8da31-515">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-515">
         - ImageCoercion</span></span><br><span data-ttu-id="8da31-516">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-516">
         - MatrixBindings</span></span><br><span data-ttu-id="8da31-517">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-517">
         - MatrixCoercion</span></span><br><span data-ttu-id="8da31-518">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-518">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8da31-519">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8da31-519">
         - PdfFile</span></span><br><span data-ttu-id="8da31-520">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-520">
         - Selection</span></span><br><span data-ttu-id="8da31-521">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-521">
         - Settings</span></span><br><span data-ttu-id="8da31-522">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-522">
         - TableBindings</span></span><br><span data-ttu-id="8da31-523">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-523">
         - TableCoercion</span></span><br><span data-ttu-id="8da31-524">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-524">
         - TextBindings</span></span><br><span data-ttu-id="8da31-525">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-525">
         - TextCoercion</span></span><br><span data-ttu-id="8da31-526">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8da31-526">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-527">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="8da31-527">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="8da31-528">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-528">- TaskPane</span></span><br><span data-ttu-id="8da31-529">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-529">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8da31-530">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-530">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="8da31-531">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8da31-531">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="8da31-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8da31-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="8da31-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="8da31-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="8da31-534">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-534">- BindingEvents</span></span><br><span data-ttu-id="8da31-535">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-535">
         - CompressedFile</span></span><br><span data-ttu-id="8da31-536">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8da31-536">
         - CustomXmlParts</span></span><br><span data-ttu-id="8da31-537">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-537">
         - DocumentEvents</span></span><br><span data-ttu-id="8da31-538">
         - File</span><span class="sxs-lookup"><span data-stu-id="8da31-538">
         - File</span></span><br><span data-ttu-id="8da31-539">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-539">
         - HtmlCoercion</span></span><br><span data-ttu-id="8da31-540">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-540">
         - ImageCoercion</span></span><br><span data-ttu-id="8da31-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-541">
         - MatrixBindings</span></span><br><span data-ttu-id="8da31-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="8da31-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8da31-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8da31-544">
         - PdfFile</span></span><br><span data-ttu-id="8da31-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-545">
         - Selection</span></span><br><span data-ttu-id="8da31-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-546">
         - Settings</span></span><br><span data-ttu-id="8da31-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-547">
         - TableBindings</span></span><br><span data-ttu-id="8da31-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-548">
         - TableCoercion</span></span><br><span data-ttu-id="8da31-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8da31-549">
         - TextBindings</span></span><br><span data-ttu-id="8da31-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-550">
         - TextCoercion</span></span><br><span data-ttu-id="8da31-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8da31-551">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="8da31-552">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="8da31-552">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8da31-553">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="8da31-553">Platform</span></span></th>
    <th><span data-ttu-id="8da31-554">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="8da31-554">Extension points</span></span></th>
    <th><span data-ttu-id="8da31-555">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="8da31-555">API requirement sets</span></span></th>
    <th><span data-ttu-id="8da31-556"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="8da31-556"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-557">Office Online</span><span class="sxs-lookup"><span data-stu-id="8da31-557">Office Online</span></span></td>
    <td> <span data-ttu-id="8da31-558">- Contenu</span><span class="sxs-lookup"><span data-stu-id="8da31-558">- Content</span></span><br><span data-ttu-id="8da31-559">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-559">
         - TaskPane</span></span><br><span data-ttu-id="8da31-560">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-560">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8da31-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8da31-562">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8da31-562">- ActiveView</span></span><br><span data-ttu-id="8da31-563">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-563">
         - CompressedFile</span></span><br><span data-ttu-id="8da31-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-564">
         - DocumentEvents</span></span><br><span data-ttu-id="8da31-565">
         - File</span><span class="sxs-lookup"><span data-stu-id="8da31-565">
         - File</span></span><br><span data-ttu-id="8da31-566">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-566">
         - ImageCoercion</span></span><br><span data-ttu-id="8da31-567">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8da31-567">
         - PdfFile</span></span><br><span data-ttu-id="8da31-568">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-568">
         - Selection</span></span><br><span data-ttu-id="8da31-569">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-569">
         - Settings</span></span><br><span data-ttu-id="8da31-570">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-570">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-571">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="8da31-571">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="8da31-572">- Contenu</span><span class="sxs-lookup"><span data-stu-id="8da31-572">- Content</span></span><br><span data-ttu-id="8da31-573">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-573">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="8da31-574">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="8da31-574">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="8da31-575">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8da31-575">- ActiveView</span></span><br><span data-ttu-id="8da31-576">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-576">
         - CompressedFile</span></span><br><span data-ttu-id="8da31-577">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-577">
         - DocumentEvents</span></span><br><span data-ttu-id="8da31-578">
         - File</span><span class="sxs-lookup"><span data-stu-id="8da31-578">
         - File</span></span><br><span data-ttu-id="8da31-579">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-579">
         - ImageCoercion</span></span><br><span data-ttu-id="8da31-580">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8da31-580">
         - PdfFile</span></span><br><span data-ttu-id="8da31-581">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-581">
         - Selection</span></span><br><span data-ttu-id="8da31-582">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-582">
         - Settings</span></span><br><span data-ttu-id="8da31-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-583">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-584">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="8da31-584">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="8da31-585">- Contenu</span><span class="sxs-lookup"><span data-stu-id="8da31-585">- Content</span></span><br><span data-ttu-id="8da31-586">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-586">
         - TaskPane</span></span><br><span data-ttu-id="8da31-587">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-587">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8da31-588">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-588">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8da31-589">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8da31-589">- ActiveView</span></span><br><span data-ttu-id="8da31-590">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-590">
         - CompressedFile</span></span><br><span data-ttu-id="8da31-591">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-591">
         - DocumentEvents</span></span><br><span data-ttu-id="8da31-592">
         - File</span><span class="sxs-lookup"><span data-stu-id="8da31-592">
         - File</span></span><br><span data-ttu-id="8da31-593">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-593">
         - ImageCoercion</span></span><br><span data-ttu-id="8da31-594">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8da31-594">
         - PdfFile</span></span><br><span data-ttu-id="8da31-595">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-595">
         - Selection</span></span><br><span data-ttu-id="8da31-596">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-596">
         - Settings</span></span><br><span data-ttu-id="8da31-597">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-597">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-598">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="8da31-598">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="8da31-599">- Contenu</span><span class="sxs-lookup"><span data-stu-id="8da31-599">- Content</span></span><br><span data-ttu-id="8da31-600">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-600">
         - TaskPane</span></span><br><span data-ttu-id="8da31-601">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-601">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8da31-602">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-602">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8da31-603">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8da31-603">- ActiveView</span></span><br><span data-ttu-id="8da31-604">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-604">
         - CompressedFile</span></span><br><span data-ttu-id="8da31-605">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-605">
         - DocumentEvents</span></span><br><span data-ttu-id="8da31-606">
         - File</span><span class="sxs-lookup"><span data-stu-id="8da31-606">
         - File</span></span><br><span data-ttu-id="8da31-607">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-607">
         - ImageCoercion</span></span><br><span data-ttu-id="8da31-608">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8da31-608">
         - PdfFile</span></span><br><span data-ttu-id="8da31-609">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-609">
         - Selection</span></span><br><span data-ttu-id="8da31-610">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-610">
         - Settings</span></span><br><span data-ttu-id="8da31-611">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-611">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-612">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="8da31-612">Office for iPad</span></span></td>
    <td> <span data-ttu-id="8da31-613">- Contenu</span><span class="sxs-lookup"><span data-stu-id="8da31-613">- Content</span></span><br><span data-ttu-id="8da31-614">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-614">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="8da31-615">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-615">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="8da31-616">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8da31-616">- ActiveView</span></span><br><span data-ttu-id="8da31-617">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-617">
         - CompressedFile</span></span><br><span data-ttu-id="8da31-618">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-618">
         - DocumentEvents</span></span><br><span data-ttu-id="8da31-619">
         - File</span><span class="sxs-lookup"><span data-stu-id="8da31-619">
         - File</span></span><br><span data-ttu-id="8da31-620">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8da31-620">
         - PdfFile</span></span><br><span data-ttu-id="8da31-621">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-621">
         - Selection</span></span><br><span data-ttu-id="8da31-622">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-622">
         - Settings</span></span><br><span data-ttu-id="8da31-623">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-623">
         - TextCoercion</span></span><br><span data-ttu-id="8da31-624">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-624">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-625">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="8da31-625">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="8da31-626">- Contenu</span><span class="sxs-lookup"><span data-stu-id="8da31-626">- Content</span></span><br><span data-ttu-id="8da31-627">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-627">
         - TaskPane</span></span><br><span data-ttu-id="8da31-628">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-628">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8da31-629">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-629">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8da31-630">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8da31-630">- ActiveView</span></span><br><span data-ttu-id="8da31-631">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-631">
         - CompressedFile</span></span><br><span data-ttu-id="8da31-632">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-632">
         - DocumentEvents</span></span><br><span data-ttu-id="8da31-633">
         - File</span><span class="sxs-lookup"><span data-stu-id="8da31-633">
         - File</span></span><br><span data-ttu-id="8da31-634">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-634">
         - ImageCoercion</span></span><br><span data-ttu-id="8da31-635">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8da31-635">
         - PdfFile</span></span><br><span data-ttu-id="8da31-636">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-636">
         - Selection</span></span><br><span data-ttu-id="8da31-637">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-637">
         - Settings</span></span><br><span data-ttu-id="8da31-638">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-638">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-639">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="8da31-639">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="8da31-640">- Contenu</span><span class="sxs-lookup"><span data-stu-id="8da31-640">- Content</span></span><br><span data-ttu-id="8da31-641">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-641">
         - TaskPane</span></span><br><span data-ttu-id="8da31-642">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-642">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8da31-643">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-643">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8da31-644">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8da31-644">- ActiveView</span></span><br><span data-ttu-id="8da31-645">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8da31-645">
         - CompressedFile</span></span><br><span data-ttu-id="8da31-646">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-646">
         - DocumentEvents</span></span><br><span data-ttu-id="8da31-647">
         - File</span><span class="sxs-lookup"><span data-stu-id="8da31-647">
         - File</span></span><br><span data-ttu-id="8da31-648">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-648">
         - ImageCoercion</span></span><br><span data-ttu-id="8da31-649">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8da31-649">
         - PdfFile</span></span><br><span data-ttu-id="8da31-650">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-650">
         - Selection</span></span><br><span data-ttu-id="8da31-651">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-651">
         - Settings</span></span><br><span data-ttu-id="8da31-652">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-652">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="8da31-653">OneNote</span><span class="sxs-lookup"><span data-stu-id="8da31-653">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8da31-654">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="8da31-654">Platform</span></span></th>
    <th><span data-ttu-id="8da31-655">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="8da31-655">Extension points</span></span></th>
    <th><span data-ttu-id="8da31-656">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="8da31-656">API requirement sets</span></span></th>
    <th><span data-ttu-id="8da31-657"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="8da31-657"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-658">Office Online</span><span class="sxs-lookup"><span data-stu-id="8da31-658">Office Online</span></span></td>
    <td> <span data-ttu-id="8da31-659">- Contenu</span><span class="sxs-lookup"><span data-stu-id="8da31-659">- Content</span></span><br><span data-ttu-id="8da31-660">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-660">
         - TaskPane</span></span><br><span data-ttu-id="8da31-661">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="8da31-661">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8da31-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="8da31-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8da31-664">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8da31-664">- DocumentEvents</span></span><br><span data-ttu-id="8da31-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="8da31-666">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-666">
         - ImageCoercion</span></span><br><span data-ttu-id="8da31-667">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="8da31-667">
         - Settings</span></span><br><span data-ttu-id="8da31-668">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-668">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="8da31-669">Projet</span><span class="sxs-lookup"><span data-stu-id="8da31-669">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8da31-670">Plateforme</span><span class="sxs-lookup"><span data-stu-id="8da31-670">Platform</span></span></th>
    <th><span data-ttu-id="8da31-671">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="8da31-671">Extension points</span></span></th>
    <th><span data-ttu-id="8da31-672">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="8da31-672">API requirement sets</span></span></th>
    <th><span data-ttu-id="8da31-673"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="8da31-673"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-674">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="8da31-674">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="8da31-675">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-675">- TaskPane</span></span></td>
    <td> <span data-ttu-id="8da31-676">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-676">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8da31-677">- Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-677">- Selection</span></span><br><span data-ttu-id="8da31-678">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-678">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-679">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="8da31-679">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="8da31-680">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-680">- TaskPane</span></span></td>
    <td> <span data-ttu-id="8da31-681">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-681">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8da31-682">- Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-682">- Selection</span></span><br><span data-ttu-id="8da31-683">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-683">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8da31-684">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="8da31-684">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="8da31-685">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="8da31-685">- TaskPane</span></span></td>
    <td> <span data-ttu-id="8da31-686">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8da31-686">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8da31-687">- Selection</span><span class="sxs-lookup"><span data-stu-id="8da31-687">- Selection</span></span><br><span data-ttu-id="8da31-688">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8da31-688">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="8da31-689">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8da31-689">See also</span></span>

- [<span data-ttu-id="8da31-690">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="8da31-690">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="8da31-691">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="8da31-691">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="8da31-692">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="8da31-692">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="8da31-693">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="8da31-693">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
