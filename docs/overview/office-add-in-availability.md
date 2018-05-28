---
title: Disponibilit? des compl?ments Office sur les plateformes et les h?tes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint et OneNote.
ms.date: 03/23/2018
ms.openlocfilehash: f50ab7e5312702eb25fbb2c8a25291c5ff5027a7
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="6e7b0-103">Disponibilit? des compl?ments Office sur les plateformes et les h?tes</span><span class="sxs-lookup"><span data-stu-id="6e7b0-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="6e7b0-104">Pour fonctionner comme pr?vu, il se peut que votre compl?ment Office d?pende d?un h?te Office sp?cifique, d?un ensemble de conditions requises, d?un membre d?API ou d?une version de l?API.</span><span class="sxs-lookup"><span data-stu-id="6e7b0-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="6e7b0-105">Les tableaux suivants contiennent la plateforme disponible, les points d?extension, les ensembles de conditions requises de l?API et les ensembles de conditions requises des API communes qui sont actuellement pris en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="6e7b0-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span> 

<span data-ttu-id="6e7b0-106">Si une cellule de tableau contient un ast?risque (\*), cela signifie que nous travaillons sur celle-ci.</span><span class="sxs-lookup"><span data-stu-id="6e7b0-106">If a table cell contains an asterisk ( \* ), that means we?re working on it.</span></span> <span data-ttu-id="6e7b0-107">Pour les ensembles de conditions requises pour Project ou Access, consultez les [ensembles de conditions requises communs ? Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="6e7b0-107">For requirement sets for Project or Access, see [Office common requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="6e7b0-p103">Le num?ro de build pour Office 2016 install? via MSI est 16.0.4266.1001. Cette version ne contient que les ensembles de conditions requises des API communes, ExcelApi 1.1 et WordApi 1.1.</span><span class="sxs-lookup"><span data-stu-id="6e7b0-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="6e7b0-110">Excel</span><span class="sxs-lookup"><span data-stu-id="6e7b0-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="6e7b0-111">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="6e7b0-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="6e7b0-112">Points d?extension</span><span class="sxs-lookup"><span data-stu-id="6e7b0-112">Extension points</span></span></th> 
    <th style="width:20%"><span data-ttu-id="6e7b0-113">Ensembles de conditions requises de l?API</span><span class="sxs-lookup"><span data-stu-id="6e7b0-113">API requirement sets</span></span></th> 
    <th style="width:40%"><span data-ttu-id="6e7b0-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="6e7b0-115">Office Online</span></span></td>
    <td> <span data-ttu-id="6e7b0-116">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="6e7b0-116">- Taskpane</span></span><br><span data-ttu-id="6e7b0-117">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="6e7b0-117">
        - Content</span></span><br><span data-ttu-id="6e7b0-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de compl?ment</a>
    </span><span class="sxs-lookup"><span data-stu-id="6e7b0-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6e7b0-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6e7b0-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6e7b0-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6e7b0-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6e7b0-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6e7b0-124">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-124">
        -BindingEvents</span></span><br><span data-ttu-id="6e7b0-125">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-125">
        -DocumentEvents</span></span><br><span data-ttu-id="6e7b0-126">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-126">
        -MatrixBindings</span></span><br><span data-ttu-id="6e7b0-127">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-127">
        -MatrixCoercion</span></span><br><span data-ttu-id="6e7b0-128">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-128">
        -TableBindings</span></span><br><span data-ttu-id="6e7b0-129">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-129">
        -TableCoercion</span></span><br><span data-ttu-id="6e7b0-130">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-130">
        -TextBindings</span></span><br><span data-ttu-id="6e7b0-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6e7b0-131">
        -CompressedFile</span></span><br><span data-ttu-id="6e7b0-132">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-132">
        - Settings</span></span><br><span data-ttu-id="6e7b0-133">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-133">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-134">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="6e7b0-134">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="6e7b0-135">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="6e7b0-135">
        - Taskpane</span></span><br><span data-ttu-id="6e7b0-136">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="6e7b0-136">
        - Content</span></span></td>
    <td>  <span data-ttu-id="6e7b0-137">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-137">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6e7b0-138">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-138">
        -BindingEvents</span></span><br><span data-ttu-id="6e7b0-139">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-139">
        -DocumentEvents</span></span><br><span data-ttu-id="6e7b0-140">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-140">
        -MatrixBindings</span></span><br><span data-ttu-id="6e7b0-141">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-141">
        -MatrixCoercion</span></span><br><span data-ttu-id="6e7b0-142">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-142">
        -TableBindings</span></span><br><span data-ttu-id="6e7b0-143">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-143">
        -TableCoercion</span></span><br><span data-ttu-id="6e7b0-144">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-144">
        -TextBindings</span></span><br><span data-ttu-id="6e7b0-145">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-145">
        - Settings</span></span><br><span data-ttu-id="6e7b0-146">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-146">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-147">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="6e7b0-147">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="6e7b0-148">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="6e7b0-148">- Taskpane</span></span><br><span data-ttu-id="6e7b0-149">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="6e7b0-149">
        - Content</span></span><br><span data-ttu-id="6e7b0-150">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de compl?ment</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-150">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="6e7b0-151">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-151">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6e7b0-152">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-152">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6e7b0-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6e7b0-154">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-154">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6e7b0-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6e7b0-156">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-156">-BindingEvents</span></span><br><span data-ttu-id="6e7b0-157">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-157">
        -DocumentEvents</span></span><br><span data-ttu-id="6e7b0-158">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-158">
        -MatrixBindings</span></span><br><span data-ttu-id="6e7b0-159">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-159">
        -MatrixCoercion</span></span><br><span data-ttu-id="6e7b0-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-160">
        -TableBindings</span></span><br><span data-ttu-id="6e7b0-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-161">
        -TableCoercion</span></span><br><span data-ttu-id="6e7b0-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-162">
        -TextBindings</span></span><br><span data-ttu-id="6e7b0-163">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-163">
        - Settings</span></span><br><span data-ttu-id="6e7b0-164">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-164">
        -TextCoercion</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-165">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="6e7b0-165">Office for iOS</span></span></td>
    <td><span data-ttu-id="6e7b0-166">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="6e7b0-166">- Taskpane</span></span><br><span data-ttu-id="6e7b0-167">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="6e7b0-167">
        - Content</span></span></td>
    <td><span data-ttu-id="6e7b0-168">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-168">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6e7b0-169">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-169">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6e7b0-170">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-170">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6e7b0-171">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-171">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6e7b0-172">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-172">-BindingEvents</span></span><br><span data-ttu-id="6e7b0-173">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-173">
        -DocumentEvents</span></span><br><span data-ttu-id="6e7b0-174">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-174">
        -MatrixBindings</span></span><br><span data-ttu-id="6e7b0-175">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-175">
        -MatrixCoercion</span></span><br><span data-ttu-id="6e7b0-176">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-176">
        -TableBindings</span></span><br><span data-ttu-id="6e7b0-177">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-177">
        -TableCoercion</span></span><br><span data-ttu-id="6e7b0-178">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-178">
        -TextBindings</span></span><br><span data-ttu-id="6e7b0-179">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-179">
        - Settings</span></span><br><span data-ttu-id="6e7b0-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-181">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="6e7b0-181">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="6e7b0-182">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="6e7b0-182">- Taskpane</span></span><br><span data-ttu-id="6e7b0-183">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="6e7b0-183">
        - Content</span></span><br><span data-ttu-id="6e7b0-184">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de compl?ment</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-184">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="6e7b0-185">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-185">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6e7b0-186">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-186">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6e7b0-187">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-187">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6e7b0-188">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-188">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6e7b0-189">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-189">-BindingEvents</span></span><br><span data-ttu-id="6e7b0-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-190">
        -DocumentEvents</span></span><br><span data-ttu-id="6e7b0-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-191">
        -MatrixBindings</span></span><br><span data-ttu-id="6e7b0-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-192">
        -MatrixCoercion</span></span><br><span data-ttu-id="6e7b0-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-193">
        -TableBindings</span></span><br><span data-ttu-id="6e7b0-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-194">
        -TableCoercion</span></span><br><span data-ttu-id="6e7b0-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-195">
        -TextBindings</span></span><br><span data-ttu-id="6e7b0-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-196">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="6e7b0-197">Outlook</span><span class="sxs-lookup"><span data-stu-id="6e7b0-197">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6e7b0-198">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="6e7b0-198">Platform</span></span></th>
    <th><span data-ttu-id="6e7b0-199">Points d?extension</span><span class="sxs-lookup"><span data-stu-id="6e7b0-199">Extension points</span></span></th> 
    <th><span data-ttu-id="6e7b0-200">Ensembles de conditions requises de l?API</span><span class="sxs-lookup"><span data-stu-id="6e7b0-200">API requirement sets</span></span></th> 
    <th><span data-ttu-id="6e7b0-201"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-201"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-202">Office Online</span><span class="sxs-lookup"><span data-stu-id="6e7b0-202">Office Online</span></span></td>
    <td> <span data-ttu-id="6e7b0-203">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="6e7b0-203">- Mail Read</span></span><br><span data-ttu-id="6e7b0-204">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="6e7b0-204">
      - Mail Compose</span></span><br><span data-ttu-id="6e7b0-205">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de compl?ment</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-205">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6e7b0-206">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-206">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6e7b0-207">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-207">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6e7b0-208">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-208">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6e7b0-209">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-209">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6e7b0-210">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-210">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6e7b0-211">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-211">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="6e7b0-212">non disponible</span><span class="sxs-lookup"><span data-stu-id="6e7b0-212">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-213">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="6e7b0-213">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="6e7b0-214">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="6e7b0-214">- Mail Read</span></span><br><span data-ttu-id="6e7b0-215">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="6e7b0-215">
      - Mail Compose</span></span><br><span data-ttu-id="6e7b0-216">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de compl?ment</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-216">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6e7b0-217">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-217">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6e7b0-218">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-218">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6e7b0-219">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-219">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6e7b0-220">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-220">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="6e7b0-221">non disponible</span><span class="sxs-lookup"><span data-stu-id="6e7b0-221">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-222">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="6e7b0-222">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="6e7b0-223">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="6e7b0-223">- Mail Read</span></span><br><span data-ttu-id="6e7b0-224">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="6e7b0-224">
      - Mail Compose</span></span><br><span data-ttu-id="6e7b0-225">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de compl?ment</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-225">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="6e7b0-226">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="6e7b0-226">
      - Modules</span></span></td>
    <td> <span data-ttu-id="6e7b0-227">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-227">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6e7b0-228">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-228">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6e7b0-229">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-229">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6e7b0-230">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-230">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6e7b0-231">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-231">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6e7b0-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="6e7b0-233">non disponible</span><span class="sxs-lookup"><span data-stu-id="6e7b0-233">not available</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-234">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="6e7b0-234">Office for iOS</span></span></td>
    <td> <span data-ttu-id="6e7b0-235">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="6e7b0-235">- Mail Read</span></span><br><span data-ttu-id="6e7b0-236">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de compl?ment</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-236">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6e7b0-237">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-237">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6e7b0-238">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-238">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6e7b0-239">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-239">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6e7b0-240">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-240">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6e7b0-241">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-241">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>    
    <td><span data-ttu-id="6e7b0-242">non disponible</span><span class="sxs-lookup"><span data-stu-id="6e7b0-242">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-243">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="6e7b0-243">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="6e7b0-244">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="6e7b0-244">- Mail Read</span></span><br><span data-ttu-id="6e7b0-245">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="6e7b0-245">
      - Mail Compose</span></span><br><span data-ttu-id="6e7b0-246">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de compl?ment</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-246">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6e7b0-247">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-247">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6e7b0-248">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-248">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6e7b0-249">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-249">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6e7b0-250">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-250">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6e7b0-251">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-251">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6e7b0-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="6e7b0-253">non disponible</span><span class="sxs-lookup"><span data-stu-id="6e7b0-253">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-254">Office pour Android</span><span class="sxs-lookup"><span data-stu-id="6e7b0-254">Office for Android</span></span></td>
    <td> <span data-ttu-id="6e7b0-255">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="6e7b0-255">- Mail Read</span></span><br><span data-ttu-id="6e7b0-256">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de compl?ment</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-256">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6e7b0-257">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-257">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6e7b0-258">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-258">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6e7b0-259">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-259">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6e7b0-260">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-260">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6e7b0-261">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-261">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="6e7b0-262">non disponible</span><span class="sxs-lookup"><span data-stu-id="6e7b0-262">not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="6e7b0-263">Word</span><span class="sxs-lookup"><span data-stu-id="6e7b0-263">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6e7b0-264">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="6e7b0-264">Platform</span></span></th>
    <th><span data-ttu-id="6e7b0-265">Points d?extension</span><span class="sxs-lookup"><span data-stu-id="6e7b0-265">Extension points</span></span></th> 
    <th><span data-ttu-id="6e7b0-266">Ensembles de conditions requises de l?API</span><span class="sxs-lookup"><span data-stu-id="6e7b0-266">API requirement sets</span></span></th> 
    <th><span data-ttu-id="6e7b0-267"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-267"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-268">Office Online</span><span class="sxs-lookup"><span data-stu-id="6e7b0-268">Office Online</span></span></td>
    <td> <span data-ttu-id="6e7b0-269">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="6e7b0-269">- Taskpane</span></span><br><span data-ttu-id="6e7b0-270">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de compl?ment</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-270">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6e7b0-271">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-271">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="6e7b0-272">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-272">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="6e7b0-273">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-273">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="6e7b0-274">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-274">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6e7b0-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-275">-BindingEvents</span></span><br><span data-ttu-id="6e7b0-276">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6e7b0-276">
         -</span></span><br><span data-ttu-id="6e7b0-277">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-277">
         -MatrixBindings</span></span><br><span data-ttu-id="6e7b0-278">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-278">
         -MatrixCoercion</span></span><br><span data-ttu-id="6e7b0-279">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-279">
         -TableBindings</span></span><br><span data-ttu-id="6e7b0-280">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-280">
         -TableCoercion</span></span><br><span data-ttu-id="6e7b0-281">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-281">
         -TextBindings</span></span><br><span data-ttu-id="6e7b0-282">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-282">
         -DocumentEvents</span></span><br><span data-ttu-id="6e7b0-283">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6e7b0-283">
         -TextFile</span></span><br><span data-ttu-id="6e7b0-284">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-284">
         -ImageCoercion</span></span><br><span data-ttu-id="6e7b0-285">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-285">
         - Settings</span></span><br><span data-ttu-id="6e7b0-286">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-286">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-287">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="6e7b0-287">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="6e7b0-288">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="6e7b0-288">- Taskpane</span></span></td>
    <td> <span data-ttu-id="6e7b0-289">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-289">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6e7b0-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-290">-BindingEvents</span></span><br><span data-ttu-id="6e7b0-291">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6e7b0-291">
         -CompressedFile</span></span><br><span data-ttu-id="6e7b0-292">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="6e7b0-292">
         -CustomXmlPart</span></span><br><span data-ttu-id="6e7b0-293">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-293">
         -DocumentEvents</span></span><br><span data-ttu-id="6e7b0-294">
         - File</span><span class="sxs-lookup"><span data-stu-id="6e7b0-294">
         - File</span></span><br><span data-ttu-id="6e7b0-295">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-295">
         -HtmlCoercion</span></span><br><span data-ttu-id="6e7b0-296">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-296">
         -ImageCoercion</span></span><br><span data-ttu-id="6e7b0-297">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-297">
         -OoxmlCoercion</span></span><br><span data-ttu-id="6e7b0-298">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-298">
         -TableBindings</span></span><br><span data-ttu-id="6e7b0-299">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-299">
         -TableCoercion</span></span><br><span data-ttu-id="6e7b0-300">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-300">
         -TextBindings</span></span><br><span data-ttu-id="6e7b0-301">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6e7b0-301">
         -TextFile</span></span><br><span data-ttu-id="6e7b0-302">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-302">
         - Settings</span></span><br><span data-ttu-id="6e7b0-303">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-303">
         -TextCoercion</span></span><br><span data-ttu-id="6e7b0-304">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-304">
         -MatrixCoercion</span></span><br><span data-ttu-id="6e7b0-305">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-305">
         - Matrix Bindings</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-306">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="6e7b0-306">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="6e7b0-307">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="6e7b0-307">- Taskpane</span></span><br><span data-ttu-id="6e7b0-308">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de compl?ment</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-308">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6e7b0-309">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-309">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="6e7b0-310">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-310">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="6e7b0-311">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-311">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="6e7b0-312">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-312">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6e7b0-313">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-313">-BindingEvents</span></span><br><span data-ttu-id="6e7b0-314">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6e7b0-314">
         -CompressedFile</span></span><br><span data-ttu-id="6e7b0-315">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="6e7b0-315">
         -CustomXmlPart</span></span><br><span data-ttu-id="6e7b0-316">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-316">
         -DocumentEvents</span></span><br><span data-ttu-id="6e7b0-317">
         - File</span><span class="sxs-lookup"><span data-stu-id="6e7b0-317">
         - File</span></span><br><span data-ttu-id="6e7b0-318">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-318">
         -HtmlCoercion</span></span><br><span data-ttu-id="6e7b0-319">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-319">
         -ImageCoercion</span></span><br><span data-ttu-id="6e7b0-320">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-320">
         -OoxmlCoercion</span></span><br><span data-ttu-id="6e7b0-321">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-321">
         -TableBindings</span></span><br><span data-ttu-id="6e7b0-322">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-322">
         -TableCoercion</span></span><br><span data-ttu-id="6e7b0-323">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-323">
         -TextBindings</span></span><br><span data-ttu-id="6e7b0-324">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6e7b0-324">
         -TextFile</span></span><br><span data-ttu-id="6e7b0-325">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-325">
         - Settings</span></span><br><span data-ttu-id="6e7b0-326">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-326">
         -TextCoercion</span></span><br><span data-ttu-id="6e7b0-327">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-327">
         -MatrixCoercion</span></span><br><span data-ttu-id="6e7b0-328">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-328">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-329">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="6e7b0-329">Office for iOS</span></span></td>
    <td> <span data-ttu-id="6e7b0-330">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="6e7b0-330">- Taskpane</span></span></td>
    <td> <span data-ttu-id="6e7b0-331">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-331">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="6e7b0-332">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-332">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="6e7b0-333">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-333">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="6e7b0-334">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="6e7b0-334">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="6e7b0-335">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-335">-BindingEvents</span></span><br><span data-ttu-id="6e7b0-336">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6e7b0-336">
         -CompressedFile</span></span><br><span data-ttu-id="6e7b0-337">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="6e7b0-337">
         -CustomXmlPart</span></span><br><span data-ttu-id="6e7b0-338">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-338">
         -DocumentEvents</span></span><br><span data-ttu-id="6e7b0-339">
         - File</span><span class="sxs-lookup"><span data-stu-id="6e7b0-339">
         - File</span></span><br><span data-ttu-id="6e7b0-340">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-340">
         -HtmlCoercion</span></span><br><span data-ttu-id="6e7b0-341">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-341">
         -ImageCoercion</span></span><br><span data-ttu-id="6e7b0-342">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-342">
         -OoxmlCoercion</span></span><br><span data-ttu-id="6e7b0-343">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-343">
         -TableBindings</span></span><br><span data-ttu-id="6e7b0-344">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-344">
         -TableCoercion</span></span><br><span data-ttu-id="6e7b0-345">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-345">
         -TextBindings</span></span><br><span data-ttu-id="6e7b0-346">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6e7b0-346">
         -TextFile</span></span><br><span data-ttu-id="6e7b0-347">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-347">
         - Settings</span></span><br><span data-ttu-id="6e7b0-348">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-348">
         -TextCoercion</span></span><br><span data-ttu-id="6e7b0-349">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-349">
         -MatrixCoercion</span></span><br><span data-ttu-id="6e7b0-350">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-350">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-351">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="6e7b0-351">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="6e7b0-352">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="6e7b0-352">- Taskpane</span></span><br><span data-ttu-id="6e7b0-353">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de compl?ment</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-353">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6e7b0-354">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-354">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="6e7b0-355">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-355">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="6e7b0-356">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-356">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="6e7b0-357">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="6e7b0-357">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="6e7b0-358">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-358">-BindingEvents</span></span><br><span data-ttu-id="6e7b0-359">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6e7b0-359">
         -CompressedFile</span></span><br><span data-ttu-id="6e7b0-360">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="6e7b0-360">
         -CustomXmlPart</span></span><br><span data-ttu-id="6e7b0-361">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-361">
         -DocumentEvents</span></span><br><span data-ttu-id="6e7b0-362">
         - File</span><span class="sxs-lookup"><span data-stu-id="6e7b0-362">
         - File</span></span><br><span data-ttu-id="6e7b0-363">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-363">
         -HtmlCoercion</span></span><br><span data-ttu-id="6e7b0-364">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-364">
         -ImageCoercion</span></span><br><span data-ttu-id="6e7b0-365">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-365">
         -OoxmlCoercion</span></span><br><span data-ttu-id="6e7b0-366">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-366">
         -TableBindings</span></span><br><span data-ttu-id="6e7b0-367">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-367">
         -TableCoercion</span></span><br><span data-ttu-id="6e7b0-368">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-368">
         -TextBindings</span></span><br><span data-ttu-id="6e7b0-369">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6e7b0-369">
         -TextFile</span></span><br><span data-ttu-id="6e7b0-370">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-370">
         - Settings</span></span><br><span data-ttu-id="6e7b0-371">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-371">
         -TextCoercion</span></span><br><span data-ttu-id="6e7b0-372">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-372">
         -MatrixCoercion</span></span><br><span data-ttu-id="6e7b0-373">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-373">
         - Matrix Bindings</span></span> </td> 
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="6e7b0-374">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="6e7b0-374">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6e7b0-375">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="6e7b0-375">Platform</span></span></th>
    <th><span data-ttu-id="6e7b0-376">Points d?extension</span><span class="sxs-lookup"><span data-stu-id="6e7b0-376">Extension points</span></span></th> 
    <th><span data-ttu-id="6e7b0-377">Ensembles de conditions requises de l?API</span><span class="sxs-lookup"><span data-stu-id="6e7b0-377">API requirement sets</span></span></th> 
    <th><span data-ttu-id="6e7b0-378"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-378"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-379">Office Online</span><span class="sxs-lookup"><span data-stu-id="6e7b0-379">Office Online</span></span></td>
    <td> <span data-ttu-id="6e7b0-380">- Contenu</span><span class="sxs-lookup"><span data-stu-id="6e7b0-380">- Content</span></span><br><span data-ttu-id="6e7b0-381">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="6e7b0-381">
         - Taskpane</span></span><br><span data-ttu-id="6e7b0-382">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de compl?ment</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-382">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6e7b0-383">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-383">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6e7b0-384">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6e7b0-384">-ActiveView</span></span><br><span data-ttu-id="6e7b0-385">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6e7b0-385">
         -CompressedFile</span></span><br><span data-ttu-id="6e7b0-386">
         - File</span><span class="sxs-lookup"><span data-stu-id="6e7b0-386">
         - File</span></span><br><span data-ttu-id="6e7b0-387">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6e7b0-387">
         - Selection</span></span><br><span data-ttu-id="6e7b0-388">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-388">
         - Settings</span></span><br><span data-ttu-id="6e7b0-389">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-389">
         -TextCoercion</span></span><br><span data-ttu-id="6e7b0-390">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-390">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-391">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="6e7b0-391">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="6e7b0-392">- Contenu</span><span class="sxs-lookup"><span data-stu-id="6e7b0-392">- Content</span></span><br><span data-ttu-id="6e7b0-393">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="6e7b0-393">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="6e7b0-394">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="6e7b0-394">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="6e7b0-395">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6e7b0-395">-ActiveView</span></span><br><span data-ttu-id="6e7b0-396">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6e7b0-396">
         -CompressedFile</span></span><br><span data-ttu-id="6e7b0-397">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-397">
         -DocumentEvents</span></span><br><span data-ttu-id="6e7b0-398">
         - File</span><span class="sxs-lookup"><span data-stu-id="6e7b0-398">
         - File</span></span><br><span data-ttu-id="6e7b0-399">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6e7b0-399">
         - Selection</span></span><br><span data-ttu-id="6e7b0-400">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-400">
         - Settings</span></span><br><span data-ttu-id="6e7b0-401">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-401">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-402">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="6e7b0-402">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="6e7b0-403">- Contenu</span><span class="sxs-lookup"><span data-stu-id="6e7b0-403">- Content</span></span><br><span data-ttu-id="6e7b0-404">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="6e7b0-404">
         - Taskpane</span></span><br><span data-ttu-id="6e7b0-405">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de compl?ment</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-405">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6e7b0-406">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-406">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6e7b0-407">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6e7b0-407">-ActiveView</span></span><br><span data-ttu-id="6e7b0-408">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6e7b0-408">
         -CompressedFile</span></span><br><span data-ttu-id="6e7b0-409">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-409">
         -DocumentEvents</span></span><br><span data-ttu-id="6e7b0-410">
         - File</span><span class="sxs-lookup"><span data-stu-id="6e7b0-410">
         - File</span></span><br><span data-ttu-id="6e7b0-411">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6e7b0-411">
         - Selection</span></span><br><span data-ttu-id="6e7b0-412">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-412">
         - Settings</span></span><br><span data-ttu-id="6e7b0-413">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-413">
         -TextCoercion</span></span><br><span data-ttu-id="6e7b0-414">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-414">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-415">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="6e7b0-415">Office for iOS</span></span></td>
    <td> <span data-ttu-id="6e7b0-416">- Contenu</span><span class="sxs-lookup"><span data-stu-id="6e7b0-416">- Content</span></span><br><span data-ttu-id="6e7b0-417">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="6e7b0-417">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="6e7b0-418">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-418">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="6e7b0-419">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6e7b0-419">-ActiveView</span></span><br><span data-ttu-id="6e7b0-420">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6e7b0-420">
         -CompressedFile</span></span><br><span data-ttu-id="6e7b0-421">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-421">
         -DocumentEvents</span></span><br><span data-ttu-id="6e7b0-422">
         - File</span><span class="sxs-lookup"><span data-stu-id="6e7b0-422">
         - File</span></span><br><span data-ttu-id="6e7b0-423">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6e7b0-423">
         - Selection</span></span><br><span data-ttu-id="6e7b0-424">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-424">
         - Settings</span></span><br><span data-ttu-id="6e7b0-425">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-425">
         -TextCoercion</span></span><br><span data-ttu-id="6e7b0-426">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-426">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-427">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="6e7b0-427">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="6e7b0-428">- Contenu</span><span class="sxs-lookup"><span data-stu-id="6e7b0-428">- Content</span></span><br><span data-ttu-id="6e7b0-429">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="6e7b0-429">
         - Taskpane</span></span><br><span data-ttu-id="6e7b0-430">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de compl?ment</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-430">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6e7b0-431">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-431">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6e7b0-432">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6e7b0-432">-ActiveView</span></span><br><span data-ttu-id="6e7b0-433">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6e7b0-433">
         -CompressedFile</span></span><br><span data-ttu-id="6e7b0-434">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-434">
         -DocumentEvents</span></span><br><span data-ttu-id="6e7b0-435">
         - File</span><span class="sxs-lookup"><span data-stu-id="6e7b0-435">
         - File</span></span><br><span data-ttu-id="6e7b0-436">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6e7b0-436">
         - Selection</span></span><br><span data-ttu-id="6e7b0-437">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-437">
         - Settings</span></span><br><span data-ttu-id="6e7b0-438">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-438">
         -TextCoercion</span></span><br><span data-ttu-id="6e7b0-439">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-439">
         -ImageCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="6e7b0-440">OneNote</span><span class="sxs-lookup"><span data-stu-id="6e7b0-440">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6e7b0-441">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="6e7b0-441">Platform</span></span></th>
    <th><span data-ttu-id="6e7b0-442">Points d?extension</span><span class="sxs-lookup"><span data-stu-id="6e7b0-442">Extension points</span></span></th> 
    <th><span data-ttu-id="6e7b0-443">Ensembles de conditions requises de l?API</span><span class="sxs-lookup"><span data-stu-id="6e7b0-443">API requirement sets</span></span></th> 
    <th><span data-ttu-id="6e7b0-444"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-444"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-445">Office Online</span><span class="sxs-lookup"><span data-stu-id="6e7b0-445">Office Online</span></span></td>
    <td> <span data-ttu-id="6e7b0-446">- Contenu</span><span class="sxs-lookup"><span data-stu-id="6e7b0-446">- Content</span></span><br><span data-ttu-id="6e7b0-447">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="6e7b0-447">
         - Taskpane</span></span><br><span data-ttu-id="6e7b0-448">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de compl?ment</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-448">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6e7b0-449">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-449">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="6e7b0-450">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6e7b0-450">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6e7b0-451">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6e7b0-451">-DocumentEvents</span></span><br><span data-ttu-id="6e7b0-452">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6e7b0-452">
         - Settings</span></span><br><span data-ttu-id="6e7b0-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-453">
         -TextCoercion</span></span><br><span data-ttu-id="6e7b0-454">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-454">
         -HtmlCoercion</span></span><br><span data-ttu-id="6e7b0-455">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6e7b0-455">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-456">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="6e7b0-456">Office 2013 for Windows</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr> 
  <tr>
    <td><span data-ttu-id="6e7b0-457">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="6e7b0-457">Office 2016 for Windows</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-458">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="6e7b0-458">Office for iOS</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6e7b0-459">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="6e7b0-459">Office 2016 for Mac</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
</table>

<br/>

<span data-ttu-id="6e7b0-460">\* = Nous y travaillons.</span><span class="sxs-lookup"><span data-stu-id="6e7b0-460">\* = We're working on it.</span></span> 

## <a name="see-also"></a><span data-ttu-id="6e7b0-461">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="6e7b0-461">See also</span></span>

- [<span data-ttu-id="6e7b0-462">Vue d?ensemble de la plateforme des compl?ments Office</span><span class="sxs-lookup"><span data-stu-id="6e7b0-462">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="6e7b0-463">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="6e7b0-463">Common API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="6e7b0-464">Ensembles de conditions requises concernant les commandes de compl?ment</span><span class="sxs-lookup"><span data-stu-id="6e7b0-464">Add-in Commands requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="6e7b0-465">R?f?rence de l?API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="6e7b0-465">JavaScript API for Office reference</span></span>](https://dev.office.com/reference/add-ins/javascript-api-for-office)

