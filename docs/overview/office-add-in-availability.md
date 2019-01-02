---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint, OneNote et Project.
ms.date: 11/07/2018
ms.openlocfilehash: 9490fca9663737e2397de159169b545e3900289f
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27458040"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="f8247-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="f8247-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="f8247-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="f8247-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="f8247-p102">Le numéro de build pour Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que les ensembles de conditions requises des API communes, ExcelApi 1.1 et WordApi 1.1.</span><span class="sxs-lookup"><span data-stu-id="f8247-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="f8247-108">Excel</span><span class="sxs-lookup"><span data-stu-id="f8247-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="f8247-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="f8247-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="f8247-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="f8247-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="f8247-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="f8247-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="f8247-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="f8247-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="f8247-113">Office Online</span></span></td>
    <td> <span data-ttu-id="f8247-114">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-114">- TaskPane</span></span><br><span data-ttu-id="f8247-115">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="f8247-115">
        - Content</span></span><br><span data-ttu-id="f8247-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="f8247-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="f8247-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f8247-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f8247-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f8247-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f8247-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f8247-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f8247-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f8247-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f8247-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f8247-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f8247-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f8247-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f8247-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f8247-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="f8247-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-126">
        - BindingEvents</span></span><br><span data-ttu-id="f8247-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-127">
        - CompressedFile</span></span><br><span data-ttu-id="f8247-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-128">
        - DocumentEvents</span></span><br><span data-ttu-id="f8247-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="f8247-129">
        - File</span></span><br><span data-ttu-id="f8247-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-130">
        - MatrixBindings</span></span><br><span data-ttu-id="f8247-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="f8247-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-132">
        - Selection</span></span><br><span data-ttu-id="f8247-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-133">
        - Settings</span></span><br><span data-ttu-id="f8247-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-134">
        - TableBindings</span></span><br><span data-ttu-id="f8247-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-135">
        - TableCoercion</span></span><br><span data-ttu-id="f8247-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-136">
        - TextBindings</span></span><br><span data-ttu-id="f8247-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-138">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="f8247-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="f8247-139">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-139">
        - TaskPane</span></span><br><span data-ttu-id="f8247-140">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="f8247-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="f8247-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="f8247-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-142">
        - BindingEvents</span></span><br><span data-ttu-id="f8247-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-143">
        - CompressedFile</span></span><br><span data-ttu-id="f8247-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-144">
        - DocumentEvents</span></span><br><span data-ttu-id="f8247-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="f8247-145">
        - File</span></span><br><span data-ttu-id="f8247-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-146">
        - ImageCoercion</span></span><br><span data-ttu-id="f8247-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-147">
        - MatrixBindings</span></span><br><span data-ttu-id="f8247-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="f8247-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-149">
        - Selection</span></span><br><span data-ttu-id="f8247-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-150">
        - Settings</span></span><br><span data-ttu-id="f8247-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-151">
        - TableBindings</span></span><br><span data-ttu-id="f8247-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-152">
        - TableCoercion</span></span><br><span data-ttu-id="f8247-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-153">
        - TextBindings</span></span><br><span data-ttu-id="f8247-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-155">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="f8247-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="f8247-156">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-156">- TaskPane</span></span><br><span data-ttu-id="f8247-157">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="f8247-157">
        - Content</span></span><br><span data-ttu-id="f8247-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="f8247-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f8247-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f8247-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f8247-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f8247-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f8247-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f8247-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f8247-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f8247-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f8247-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f8247-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f8247-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f8247-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f8247-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="f8247-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-168">- BindingEvents</span></span><br><span data-ttu-id="f8247-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-169">
        - CompressedFile</span></span><br><span data-ttu-id="f8247-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-170">
        - DocumentEvents</span></span><br><span data-ttu-id="f8247-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="f8247-171">
        - File</span></span><br><span data-ttu-id="f8247-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-172">
        - ImageCoercion</span></span><br><span data-ttu-id="f8247-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-173">
        - MatrixBindings</span></span><br><span data-ttu-id="f8247-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-174">
        - MatrixCoercion</span></span><br><span data-ttu-id="f8247-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-175">
        - Selection</span></span><br><span data-ttu-id="f8247-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-176">
        - Settings</span></span><br><span data-ttu-id="f8247-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-177">
        - TableBindings</span></span><br><span data-ttu-id="f8247-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-178">
        - TableCoercion</span></span><br><span data-ttu-id="f8247-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-179">
        - TextBindings</span></span><br><span data-ttu-id="f8247-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-180">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-181">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="f8247-181">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="f8247-182">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-182">- TaskPane</span></span><br><span data-ttu-id="f8247-183">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="f8247-183">
        - Content</span></span><br><span data-ttu-id="f8247-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="f8247-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f8247-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f8247-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f8247-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f8247-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f8247-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f8247-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f8247-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f8247-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f8247-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f8247-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f8247-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f8247-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f8247-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="f8247-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-194">- BindingEvents</span></span><br><span data-ttu-id="f8247-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-195">
        - CompressedFile</span></span><br><span data-ttu-id="f8247-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-196">
        - DocumentEvents</span></span><br><span data-ttu-id="f8247-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="f8247-197">
        - File</span></span><br><span data-ttu-id="f8247-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-198">
        - ImageCoercion</span></span><br><span data-ttu-id="f8247-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-199">
        - MatrixBindings</span></span><br><span data-ttu-id="f8247-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="f8247-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-201">
        - Selection</span></span><br><span data-ttu-id="f8247-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-202">
        - Settings</span></span><br><span data-ttu-id="f8247-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-203">
        - TableBindings</span></span><br><span data-ttu-id="f8247-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-204">
        - TableCoercion</span></span><br><span data-ttu-id="f8247-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-205">
        - TextBindings</span></span><br><span data-ttu-id="f8247-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-207">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="f8247-207">Office for iPad</span></span></td>
    <td><span data-ttu-id="f8247-208">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-208">- TaskPane</span></span><br><span data-ttu-id="f8247-209">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="f8247-209">
        - Content</span></span></td>
    <td><span data-ttu-id="f8247-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f8247-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f8247-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f8247-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f8247-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f8247-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f8247-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f8247-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f8247-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f8247-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f8247-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f8247-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f8247-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f8247-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="f8247-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-219">- BindingEvents</span></span><br><span data-ttu-id="f8247-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-220">
        - CompressedFile</span></span><br><span data-ttu-id="f8247-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-221">
        - DocumentEvents</span></span><br><span data-ttu-id="f8247-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="f8247-222">
        - File</span></span><br><span data-ttu-id="f8247-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-223">
        - ImageCoercion</span></span><br><span data-ttu-id="f8247-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-224">
        - MatrixBindings</span></span><br><span data-ttu-id="f8247-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="f8247-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-226">
        - Selection</span></span><br><span data-ttu-id="f8247-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-227">
        - Settings</span></span><br><span data-ttu-id="f8247-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-228">
        - TableBindings</span></span><br><span data-ttu-id="f8247-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-229">
        - TableCoercion</span></span><br><span data-ttu-id="f8247-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-230">
        - TextBindings</span></span><br><span data-ttu-id="f8247-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-232">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="f8247-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="f8247-233">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-233">- TaskPane</span></span><br><span data-ttu-id="f8247-234">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="f8247-234">
        - Content</span></span><br><span data-ttu-id="f8247-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="f8247-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f8247-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f8247-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f8247-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f8247-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f8247-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f8247-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f8247-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f8247-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f8247-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f8247-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f8247-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f8247-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f8247-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="f8247-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-245">- BindingEvents</span></span><br><span data-ttu-id="f8247-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-246">
        - CompressedFile</span></span><br><span data-ttu-id="f8247-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-247">
        - DocumentEvents</span></span><br><span data-ttu-id="f8247-248">
        - File</span><span class="sxs-lookup"><span data-stu-id="f8247-248">
        - File</span></span><br><span data-ttu-id="f8247-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-249">
        - ImageCoercion</span></span><br><span data-ttu-id="f8247-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-250">
        - MatrixBindings</span></span><br><span data-ttu-id="f8247-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-251">
        - MatrixCoercion</span></span><br><span data-ttu-id="f8247-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f8247-252">
        - PdfFile</span></span><br><span data-ttu-id="f8247-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-253">
        - Selection</span></span><br><span data-ttu-id="f8247-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-254">
        - Settings</span></span><br><span data-ttu-id="f8247-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-255">
        - TableBindings</span></span><br><span data-ttu-id="f8247-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-256">
        - TableCoercion</span></span><br><span data-ttu-id="f8247-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-257">
        - TextBindings</span></span><br><span data-ttu-id="f8247-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-259">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="f8247-259">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="f8247-260">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-260">- TaskPane</span></span><br><span data-ttu-id="f8247-261">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="f8247-261">
        - Content</span></span><br><span data-ttu-id="f8247-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="f8247-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f8247-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f8247-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f8247-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f8247-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f8247-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f8247-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f8247-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f8247-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f8247-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f8247-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f8247-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f8247-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f8247-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="f8247-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-272">- BindingEvents</span></span><br><span data-ttu-id="f8247-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-273">
        - CompressedFile</span></span><br><span data-ttu-id="f8247-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-274">
        - DocumentEvents</span></span><br><span data-ttu-id="f8247-275">
        - File</span><span class="sxs-lookup"><span data-stu-id="f8247-275">
        - File</span></span><br><span data-ttu-id="f8247-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-276">
        - ImageCoercion</span></span><br><span data-ttu-id="f8247-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-277">
        - MatrixBindings</span></span><br><span data-ttu-id="f8247-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-278">
        - MatrixCoercion</span></span><br><span data-ttu-id="f8247-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f8247-279">
        - PdfFile</span></span><br><span data-ttu-id="f8247-280">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-280">
        - Selection</span></span><br><span data-ttu-id="f8247-281">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-281">
        - Settings</span></span><br><span data-ttu-id="f8247-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-282">
        - TableBindings</span></span><br><span data-ttu-id="f8247-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-283">
        - TableCoercion</span></span><br><span data-ttu-id="f8247-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-284">
        - TextBindings</span></span><br><span data-ttu-id="f8247-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-285">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="f8247-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="f8247-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f8247-287">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="f8247-287">Platform</span></span></th>
    <th><span data-ttu-id="f8247-288">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="f8247-288">Extension points</span></span></th>
    <th><span data-ttu-id="f8247-289">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="f8247-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="f8247-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="f8247-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="f8247-291">Office Online</span></span></td>
    <td> <span data-ttu-id="f8247-292">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f8247-292">- Mail Read</span></span><br><span data-ttu-id="f8247-293">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="f8247-293">
      - Mail Compose</span></span><br><span data-ttu-id="f8247-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f8247-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f8247-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f8247-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f8247-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f8247-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f8247-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f8247-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f8247-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f8247-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="f8247-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f8247-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="f8247-302">Non disponible</span><span class="sxs-lookup"><span data-stu-id="f8247-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-303">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="f8247-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="f8247-304">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f8247-304">- Mail Read</span></span><br><span data-ttu-id="f8247-305">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="f8247-305">
      - Mail Compose</span></span><br><span data-ttu-id="f8247-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f8247-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f8247-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f8247-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f8247-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f8247-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="f8247-311">Non disponible</span><span class="sxs-lookup"><span data-stu-id="f8247-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-312">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="f8247-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="f8247-313">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f8247-313">- Mail Read</span></span><br><span data-ttu-id="f8247-314">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="f8247-314">
      - Mail Compose</span></span><br><span data-ttu-id="f8247-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="f8247-316">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="f8247-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="f8247-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f8247-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f8247-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f8247-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f8247-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f8247-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f8247-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f8247-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f8247-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="f8247-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f8247-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="f8247-324">Non disponible</span><span class="sxs-lookup"><span data-stu-id="f8247-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-325">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="f8247-325">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="f8247-326">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f8247-326">- Mail Read</span></span><br><span data-ttu-id="f8247-327">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="f8247-327">
      - Mail Compose</span></span><br><span data-ttu-id="f8247-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="f8247-329">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="f8247-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="f8247-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f8247-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f8247-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f8247-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f8247-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f8247-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f8247-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f8247-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f8247-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="f8247-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f8247-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="f8247-337">Non disponible</span><span class="sxs-lookup"><span data-stu-id="f8247-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-338">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="f8247-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="f8247-339">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f8247-339">- Mail Read</span></span><br><span data-ttu-id="f8247-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f8247-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f8247-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f8247-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f8247-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f8247-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f8247-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f8247-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="f8247-346">Non disponible</span><span class="sxs-lookup"><span data-stu-id="f8247-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-347">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="f8247-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="f8247-348">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f8247-348">- Mail Read</span></span><br><span data-ttu-id="f8247-349">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="f8247-349">
      - Mail Compose</span></span><br><span data-ttu-id="f8247-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f8247-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f8247-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f8247-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f8247-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f8247-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f8247-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f8247-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f8247-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f8247-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="f8247-357">Non disponible</span><span class="sxs-lookup"><span data-stu-id="f8247-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-358">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="f8247-358">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="f8247-359">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f8247-359">- Mail Read</span></span><br><span data-ttu-id="f8247-360">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="f8247-360">
      - Mail Compose</span></span><br><span data-ttu-id="f8247-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f8247-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f8247-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f8247-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f8247-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f8247-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f8247-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f8247-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f8247-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f8247-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="f8247-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f8247-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="f8247-369">Non disponible</span><span class="sxs-lookup"><span data-stu-id="f8247-369">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-370">Office pour Android</span><span class="sxs-lookup"><span data-stu-id="f8247-370">Office for Android</span></span></td>
    <td> <span data-ttu-id="f8247-371">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="f8247-371">- Mail Read</span></span><br><span data-ttu-id="f8247-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f8247-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f8247-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f8247-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f8247-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f8247-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f8247-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f8247-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="f8247-378">Non disponible</span><span class="sxs-lookup"><span data-stu-id="f8247-378">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="f8247-379">Word</span><span class="sxs-lookup"><span data-stu-id="f8247-379">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f8247-380">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="f8247-380">Platform</span></span></th>
    <th><span data-ttu-id="f8247-381">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="f8247-381">Extension points</span></span></th>
    <th><span data-ttu-id="f8247-382">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="f8247-382">API requirement sets</span></span></th>
    <th><span data-ttu-id="f8247-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="f8247-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-384">Office Online</span><span class="sxs-lookup"><span data-stu-id="f8247-384">Office Online</span></span></td>
    <td> <span data-ttu-id="f8247-385">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-385">- TaskPane</span></span><br><span data-ttu-id="f8247-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f8247-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="f8247-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="f8247-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="f8247-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f8247-391">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-391">- BindingEvents</span></span><br><span data-ttu-id="f8247-392">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f8247-392">
         - CustomXmlParts</span></span><br><span data-ttu-id="f8247-393">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-393">
         - DocumentEvents</span></span><br><span data-ttu-id="f8247-394">
         - File</span><span class="sxs-lookup"><span data-stu-id="f8247-394">
         - File</span></span><br><span data-ttu-id="f8247-395">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-395">
         - HtmlCoercion</span></span><br><span data-ttu-id="f8247-396">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-396">
         - ImageCoercion</span></span><br><span data-ttu-id="f8247-397">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-397">
         - MatrixBindings</span></span><br><span data-ttu-id="f8247-398">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-398">
         - MatrixCoercion</span></span><br><span data-ttu-id="f8247-399">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-399">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f8247-400">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f8247-400">
         - PdfFile</span></span><br><span data-ttu-id="f8247-401">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-401">
         - Selection</span></span><br><span data-ttu-id="f8247-402">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-402">
         - Settings</span></span><br><span data-ttu-id="f8247-403">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-403">
         - TableBindings</span></span><br><span data-ttu-id="f8247-404">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-404">
         - TableCoercion</span></span><br><span data-ttu-id="f8247-405">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-405">
         - TextBindings</span></span><br><span data-ttu-id="f8247-406">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-406">
         - TextCoercion</span></span><br><span data-ttu-id="f8247-407">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f8247-407">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-408">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="f8247-408">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="f8247-409">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-409">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f8247-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f8247-411">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-411">- BindingEvents</span></span><br><span data-ttu-id="f8247-412">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-412">
         - CompressedFile</span></span><br><span data-ttu-id="f8247-413">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f8247-413">
         - CustomXmlParts</span></span><br><span data-ttu-id="f8247-414">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-414">
         - DocumentEvents</span></span><br><span data-ttu-id="f8247-415">
         - File</span><span class="sxs-lookup"><span data-stu-id="f8247-415">
         - File</span></span><br><span data-ttu-id="f8247-416">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-416">
         - HtmlCoercion</span></span><br><span data-ttu-id="f8247-417">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-417">
         - ImageCoercion</span></span><br><span data-ttu-id="f8247-418">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-418">
         - MatrixBindings</span></span><br><span data-ttu-id="f8247-419">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-419">
         - MatrixCoercion</span></span><br><span data-ttu-id="f8247-420">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-420">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f8247-421">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f8247-421">
         - PdfFile</span></span><br><span data-ttu-id="f8247-422">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-422">
         - Selection</span></span><br><span data-ttu-id="f8247-423">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-423">
         - Settings</span></span><br><span data-ttu-id="f8247-424">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-424">
         - TableBindings</span></span><br><span data-ttu-id="f8247-425">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-425">
         - TableCoercion</span></span><br><span data-ttu-id="f8247-426">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-426">
         - TextBindings</span></span><br><span data-ttu-id="f8247-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-427">
         - TextCoercion</span></span><br><span data-ttu-id="f8247-428">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f8247-428">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-429">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="f8247-429">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="f8247-430">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-430">- TaskPane</span></span><br><span data-ttu-id="f8247-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f8247-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="f8247-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="f8247-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="f8247-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f8247-436">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-436">- BindingEvents</span></span><br><span data-ttu-id="f8247-437">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-437">
         - CompressedFile</span></span><br><span data-ttu-id="f8247-438">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f8247-438">
         - CustomXmlParts</span></span><br><span data-ttu-id="f8247-439">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-439">
         - DocumentEvents</span></span><br><span data-ttu-id="f8247-440">
         - File</span><span class="sxs-lookup"><span data-stu-id="f8247-440">
         - File</span></span><br><span data-ttu-id="f8247-441">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-441">
         - HtmlCoercion</span></span><br><span data-ttu-id="f8247-442">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-442">
         - ImageCoercion</span></span><br><span data-ttu-id="f8247-443">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-443">
         - MatrixBindings</span></span><br><span data-ttu-id="f8247-444">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-444">
         - MatrixCoercion</span></span><br><span data-ttu-id="f8247-445">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-445">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f8247-446">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f8247-446">
         - PdfFile</span></span><br><span data-ttu-id="f8247-447">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-447">
         - Selection</span></span><br><span data-ttu-id="f8247-448">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-448">
         - Settings</span></span><br><span data-ttu-id="f8247-449">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-449">
         - TableBindings</span></span><br><span data-ttu-id="f8247-450">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-450">
         - TableCoercion</span></span><br><span data-ttu-id="f8247-451">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-451">
         - TextBindings</span></span><br><span data-ttu-id="f8247-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-452">
         - TextCoercion</span></span><br><span data-ttu-id="f8247-453">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f8247-453">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-454">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="f8247-454">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="f8247-455">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-455">- TaskPane</span></span><br><span data-ttu-id="f8247-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f8247-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="f8247-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="f8247-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="f8247-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f8247-461">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-461">- BindingEvents</span></span><br><span data-ttu-id="f8247-462">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-462">
         - CompressedFile</span></span><br><span data-ttu-id="f8247-463">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f8247-463">
         - CustomXmlParts</span></span><br><span data-ttu-id="f8247-464">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-464">
         - DocumentEvents</span></span><br><span data-ttu-id="f8247-465">
         - File</span><span class="sxs-lookup"><span data-stu-id="f8247-465">
         - File</span></span><br><span data-ttu-id="f8247-466">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-466">
         - HtmlCoercion</span></span><br><span data-ttu-id="f8247-467">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-467">
         - ImageCoercion</span></span><br><span data-ttu-id="f8247-468">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-468">
         - MatrixBindings</span></span><br><span data-ttu-id="f8247-469">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-469">
         - MatrixCoercion</span></span><br><span data-ttu-id="f8247-470">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-470">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f8247-471">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f8247-471">
         - PdfFile</span></span><br><span data-ttu-id="f8247-472">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-472">
         - Selection</span></span><br><span data-ttu-id="f8247-473">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-473">
         - Settings</span></span><br><span data-ttu-id="f8247-474">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-474">
         - TableBindings</span></span><br><span data-ttu-id="f8247-475">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-475">
         - TableCoercion</span></span><br><span data-ttu-id="f8247-476">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-476">
         - TextBindings</span></span><br><span data-ttu-id="f8247-477">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-477">
         - TextCoercion</span></span><br><span data-ttu-id="f8247-478">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f8247-478">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-479">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="f8247-479">Office for iPad</span></span></td>
    <td> <span data-ttu-id="f8247-480">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-480">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f8247-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="f8247-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="f8247-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="f8247-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="f8247-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="f8247-485">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-485">- BindingEvents</span></span><br><span data-ttu-id="f8247-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-486">
         - CompressedFile</span></span><br><span data-ttu-id="f8247-487">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f8247-487">
         - CustomXmlParts</span></span><br><span data-ttu-id="f8247-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-488">
         - DocumentEvents</span></span><br><span data-ttu-id="f8247-489">
         - File</span><span class="sxs-lookup"><span data-stu-id="f8247-489">
         - File</span></span><br><span data-ttu-id="f8247-490">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-490">
         - HtmlCoercion</span></span><br><span data-ttu-id="f8247-491">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-491">
         - ImageCoercion</span></span><br><span data-ttu-id="f8247-492">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-492">
         - MatrixBindings</span></span><br><span data-ttu-id="f8247-493">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-493">
         - MatrixCoercion</span></span><br><span data-ttu-id="f8247-494">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-494">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f8247-495">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f8247-495">
         - PdfFile</span></span><br><span data-ttu-id="f8247-496">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-496">
         - Selection</span></span><br><span data-ttu-id="f8247-497">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-497">
         - Settings</span></span><br><span data-ttu-id="f8247-498">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-498">
         - TableBindings</span></span><br><span data-ttu-id="f8247-499">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-499">
         - TableCoercion</span></span><br><span data-ttu-id="f8247-500">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-500">
         - TextBindings</span></span><br><span data-ttu-id="f8247-501">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-501">
         - TextCoercion</span></span><br><span data-ttu-id="f8247-502">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f8247-502">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-503">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="f8247-503">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="f8247-504">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-504">- TaskPane</span></span><br><span data-ttu-id="f8247-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f8247-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="f8247-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="f8247-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="f8247-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="f8247-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="f8247-510">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-510">- BindingEvents</span></span><br><span data-ttu-id="f8247-511">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-511">
         - CompressedFile</span></span><br><span data-ttu-id="f8247-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f8247-512">
         - CustomXmlParts</span></span><br><span data-ttu-id="f8247-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-513">
         - DocumentEvents</span></span><br><span data-ttu-id="f8247-514">
         - File</span><span class="sxs-lookup"><span data-stu-id="f8247-514">
         - File</span></span><br><span data-ttu-id="f8247-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-515">
         - HtmlCoercion</span></span><br><span data-ttu-id="f8247-516">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-516">
         - ImageCoercion</span></span><br><span data-ttu-id="f8247-517">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-517">
         - MatrixBindings</span></span><br><span data-ttu-id="f8247-518">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-518">
         - MatrixCoercion</span></span><br><span data-ttu-id="f8247-519">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-519">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f8247-520">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f8247-520">
         - PdfFile</span></span><br><span data-ttu-id="f8247-521">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-521">
         - Selection</span></span><br><span data-ttu-id="f8247-522">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-522">
         - Settings</span></span><br><span data-ttu-id="f8247-523">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-523">
         - TableBindings</span></span><br><span data-ttu-id="f8247-524">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-524">
         - TableCoercion</span></span><br><span data-ttu-id="f8247-525">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-525">
         - TextBindings</span></span><br><span data-ttu-id="f8247-526">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-526">
         - TextCoercion</span></span><br><span data-ttu-id="f8247-527">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f8247-527">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-528">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="f8247-528">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="f8247-529">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-529">- TaskPane</span></span><br><span data-ttu-id="f8247-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f8247-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="f8247-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f8247-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="f8247-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f8247-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="f8247-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="f8247-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="f8247-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-535">- BindingEvents</span></span><br><span data-ttu-id="f8247-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-536">
         - CompressedFile</span></span><br><span data-ttu-id="f8247-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f8247-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="f8247-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-538">
         - DocumentEvents</span></span><br><span data-ttu-id="f8247-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="f8247-539">
         - File</span></span><br><span data-ttu-id="f8247-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="f8247-541">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-541">
         - ImageCoercion</span></span><br><span data-ttu-id="f8247-542">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-542">
         - MatrixBindings</span></span><br><span data-ttu-id="f8247-543">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-543">
         - MatrixCoercion</span></span><br><span data-ttu-id="f8247-544">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-544">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f8247-545">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f8247-545">
         - PdfFile</span></span><br><span data-ttu-id="f8247-546">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-546">
         - Selection</span></span><br><span data-ttu-id="f8247-547">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-547">
         - Settings</span></span><br><span data-ttu-id="f8247-548">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-548">
         - TableBindings</span></span><br><span data-ttu-id="f8247-549">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-549">
         - TableCoercion</span></span><br><span data-ttu-id="f8247-550">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f8247-550">
         - TextBindings</span></span><br><span data-ttu-id="f8247-551">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-551">
         - TextCoercion</span></span><br><span data-ttu-id="f8247-552">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f8247-552">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="f8247-553">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="f8247-553">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f8247-554">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="f8247-554">Platform</span></span></th>
    <th><span data-ttu-id="f8247-555">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="f8247-555">Extension points</span></span></th>
    <th><span data-ttu-id="f8247-556">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="f8247-556">API requirement sets</span></span></th>
    <th><span data-ttu-id="f8247-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="f8247-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-558">Office Online</span><span class="sxs-lookup"><span data-stu-id="f8247-558">Office Online</span></span></td>
    <td> <span data-ttu-id="f8247-559">- Contenu</span><span class="sxs-lookup"><span data-stu-id="f8247-559">- Content</span></span><br><span data-ttu-id="f8247-560">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-560">
         - TaskPane</span></span><br><span data-ttu-id="f8247-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f8247-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f8247-563">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f8247-563">- ActiveView</span></span><br><span data-ttu-id="f8247-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-564">
         - CompressedFile</span></span><br><span data-ttu-id="f8247-565">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-565">
         - DocumentEvents</span></span><br><span data-ttu-id="f8247-566">
         - File</span><span class="sxs-lookup"><span data-stu-id="f8247-566">
         - File</span></span><br><span data-ttu-id="f8247-567">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-567">
         - ImageCoercion</span></span><br><span data-ttu-id="f8247-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f8247-568">
         - PdfFile</span></span><br><span data-ttu-id="f8247-569">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-569">
         - Selection</span></span><br><span data-ttu-id="f8247-570">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-570">
         - Settings</span></span><br><span data-ttu-id="f8247-571">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-571">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-572">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="f8247-572">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="f8247-573">- Contenu</span><span class="sxs-lookup"><span data-stu-id="f8247-573">- Content</span></span><br><span data-ttu-id="f8247-574">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-574">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="f8247-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="f8247-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="f8247-576">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f8247-576">- ActiveView</span></span><br><span data-ttu-id="f8247-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-577">
         - CompressedFile</span></span><br><span data-ttu-id="f8247-578">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-578">
         - DocumentEvents</span></span><br><span data-ttu-id="f8247-579">
         - File</span><span class="sxs-lookup"><span data-stu-id="f8247-579">
         - File</span></span><br><span data-ttu-id="f8247-580">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-580">
         - ImageCoercion</span></span><br><span data-ttu-id="f8247-581">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f8247-581">
         - PdfFile</span></span><br><span data-ttu-id="f8247-582">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-582">
         - Selection</span></span><br><span data-ttu-id="f8247-583">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-583">
         - Settings</span></span><br><span data-ttu-id="f8247-584">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-584">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-585">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="f8247-585">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="f8247-586">- Contenu</span><span class="sxs-lookup"><span data-stu-id="f8247-586">- Content</span></span><br><span data-ttu-id="f8247-587">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-587">
         - TaskPane</span></span><br><span data-ttu-id="f8247-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f8247-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f8247-590">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f8247-590">- ActiveView</span></span><br><span data-ttu-id="f8247-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-591">
         - CompressedFile</span></span><br><span data-ttu-id="f8247-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-592">
         - DocumentEvents</span></span><br><span data-ttu-id="f8247-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="f8247-593">
         - File</span></span><br><span data-ttu-id="f8247-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-594">
         - ImageCoercion</span></span><br><span data-ttu-id="f8247-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f8247-595">
         - PdfFile</span></span><br><span data-ttu-id="f8247-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-596">
         - Selection</span></span><br><span data-ttu-id="f8247-597">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-597">
         - Settings</span></span><br><span data-ttu-id="f8247-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-598">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-599">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="f8247-599">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="f8247-600">- Contenu</span><span class="sxs-lookup"><span data-stu-id="f8247-600">- Content</span></span><br><span data-ttu-id="f8247-601">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-601">
         - TaskPane</span></span><br><span data-ttu-id="f8247-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f8247-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f8247-604">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f8247-604">- ActiveView</span></span><br><span data-ttu-id="f8247-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-605">
         - CompressedFile</span></span><br><span data-ttu-id="f8247-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-606">
         - DocumentEvents</span></span><br><span data-ttu-id="f8247-607">
         - File</span><span class="sxs-lookup"><span data-stu-id="f8247-607">
         - File</span></span><br><span data-ttu-id="f8247-608">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-608">
         - ImageCoercion</span></span><br><span data-ttu-id="f8247-609">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f8247-609">
         - PdfFile</span></span><br><span data-ttu-id="f8247-610">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-610">
         - Selection</span></span><br><span data-ttu-id="f8247-611">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-611">
         - Settings</span></span><br><span data-ttu-id="f8247-612">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-612">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-613">Office pour iPad</span><span class="sxs-lookup"><span data-stu-id="f8247-613">Office for iPad</span></span></td>
    <td> <span data-ttu-id="f8247-614">- Contenu</span><span class="sxs-lookup"><span data-stu-id="f8247-614">- Content</span></span><br><span data-ttu-id="f8247-615">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-615">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="f8247-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="f8247-617">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f8247-617">- ActiveView</span></span><br><span data-ttu-id="f8247-618">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-618">
         - CompressedFile</span></span><br><span data-ttu-id="f8247-619">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-619">
         - DocumentEvents</span></span><br><span data-ttu-id="f8247-620">
         - File</span><span class="sxs-lookup"><span data-stu-id="f8247-620">
         - File</span></span><br><span data-ttu-id="f8247-621">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f8247-621">
         - PdfFile</span></span><br><span data-ttu-id="f8247-622">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-622">
         - Selection</span></span><br><span data-ttu-id="f8247-623">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-623">
         - Settings</span></span><br><span data-ttu-id="f8247-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-624">
         - TextCoercion</span></span><br><span data-ttu-id="f8247-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-625">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-626">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="f8247-626">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="f8247-627">- Contenu</span><span class="sxs-lookup"><span data-stu-id="f8247-627">- Content</span></span><br><span data-ttu-id="f8247-628">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-628">
         - TaskPane</span></span><br><span data-ttu-id="f8247-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f8247-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f8247-631">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f8247-631">- ActiveView</span></span><br><span data-ttu-id="f8247-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-632">
         - CompressedFile</span></span><br><span data-ttu-id="f8247-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-633">
         - DocumentEvents</span></span><br><span data-ttu-id="f8247-634">
         - File</span><span class="sxs-lookup"><span data-stu-id="f8247-634">
         - File</span></span><br><span data-ttu-id="f8247-635">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-635">
         - ImageCoercion</span></span><br><span data-ttu-id="f8247-636">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f8247-636">
         - PdfFile</span></span><br><span data-ttu-id="f8247-637">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-637">
         - Selection</span></span><br><span data-ttu-id="f8247-638">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-638">
         - Settings</span></span><br><span data-ttu-id="f8247-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-639">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-640">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="f8247-640">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="f8247-641">- Contenu</span><span class="sxs-lookup"><span data-stu-id="f8247-641">- Content</span></span><br><span data-ttu-id="f8247-642">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-642">
         - TaskPane</span></span><br><span data-ttu-id="f8247-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f8247-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f8247-645">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f8247-645">- ActiveView</span></span><br><span data-ttu-id="f8247-646">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f8247-646">
         - CompressedFile</span></span><br><span data-ttu-id="f8247-647">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-647">
         - DocumentEvents</span></span><br><span data-ttu-id="f8247-648">
         - File</span><span class="sxs-lookup"><span data-stu-id="f8247-648">
         - File</span></span><br><span data-ttu-id="f8247-649">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-649">
         - ImageCoercion</span></span><br><span data-ttu-id="f8247-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f8247-650">
         - PdfFile</span></span><br><span data-ttu-id="f8247-651">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-651">
         - Selection</span></span><br><span data-ttu-id="f8247-652">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-652">
         - Settings</span></span><br><span data-ttu-id="f8247-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-653">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="f8247-654">OneNote</span><span class="sxs-lookup"><span data-stu-id="f8247-654">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f8247-655">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="f8247-655">Platform</span></span></th>
    <th><span data-ttu-id="f8247-656">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="f8247-656">Extension points</span></span></th>
    <th><span data-ttu-id="f8247-657">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="f8247-657">API requirement sets</span></span></th>
    <th><span data-ttu-id="f8247-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="f8247-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-659">Office Online</span><span class="sxs-lookup"><span data-stu-id="f8247-659">Office Online</span></span></td>
    <td> <span data-ttu-id="f8247-660">- Contenu</span><span class="sxs-lookup"><span data-stu-id="f8247-660">- Content</span></span><br><span data-ttu-id="f8247-661">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-661">
         - TaskPane</span></span><br><span data-ttu-id="f8247-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="f8247-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f8247-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="f8247-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f8247-665">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f8247-665">- DocumentEvents</span></span><br><span data-ttu-id="f8247-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="f8247-667">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-667">
         - ImageCoercion</span></span><br><span data-ttu-id="f8247-668">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f8247-668">
         - Settings</span></span><br><span data-ttu-id="f8247-669">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-669">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="f8247-670">Projet</span><span class="sxs-lookup"><span data-stu-id="f8247-670">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f8247-671">Plateforme</span><span class="sxs-lookup"><span data-stu-id="f8247-671">Platform</span></span></th>
    <th><span data-ttu-id="f8247-672">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="f8247-672">Extension points</span></span></th>
    <th><span data-ttu-id="f8247-673">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="f8247-673">API requirement sets</span></span></th>
    <th><span data-ttu-id="f8247-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="f8247-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-675">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="f8247-675">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="f8247-676">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-676">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f8247-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f8247-678">- Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-678">- Selection</span></span><br><span data-ttu-id="f8247-679">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-679">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-680">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="f8247-680">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="f8247-681">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-681">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f8247-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f8247-683">- Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-683">- Selection</span></span><br><span data-ttu-id="f8247-684">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-684">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f8247-685">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="f8247-685">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="f8247-686">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="f8247-686">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f8247-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f8247-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f8247-688">- Selection</span><span class="sxs-lookup"><span data-stu-id="f8247-688">- Selection</span></span><br><span data-ttu-id="f8247-689">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f8247-689">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="f8247-690">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f8247-690">See also</span></span>

- [<span data-ttu-id="f8247-691">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="f8247-691">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="f8247-692">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="f8247-692">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="f8247-693">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="f8247-693">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="f8247-694">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="f8247-694">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
