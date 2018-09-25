---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint et OneNote.
ms.date: 09/24/2018
ms.openlocfilehash: b06602e35ec906866ad16d667036a4cbaff2d89e
ms.sourcegitcommit: 8ce9a8d7f41d96879c39cc5527a3007dff25bee8
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/24/2018
ms.locfileid: "24985822"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="98fee-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="98fee-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="98fee-104">Pour fonctionner comme prévu, il se peut que votre complément Office dépende d’un hôte Office spécifique, d’un ensemble de conditions requises, d’un membre d’API ou d’une version de l’API.</span><span class="sxs-lookup"><span data-stu-id="98fee-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="98fee-105">Les tableaux suivants contiennent la plateforme disponible, les points d’extension, les ensembles de conditions requises de l’API et les ensembles de conditions requises des API communes qui sont actuellement pris en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="98fee-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="98fee-106">Si une cellule de tableau contient un astérisque (\*), cela signifie que nous travaillons sur celle-ci.</span><span class="sxs-lookup"><span data-stu-id="98fee-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="98fee-107">Pour les ensembles de conditions requises pour Project ou Access, consultez les [ensembles de conditions requises communs à Office](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="98fee-107">For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="98fee-p103">Le numéro de build pour Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que les ensembles de conditions requises des API communes, ExcelApi 1.1 et WordApi 1.1.</span><span class="sxs-lookup"><span data-stu-id="98fee-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="98fee-110">Excel</span><span class="sxs-lookup"><span data-stu-id="98fee-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="98fee-111">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="98fee-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="98fee-112">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="98fee-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="98fee-113">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="98fee-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="98fee-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="98fee-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="98fee-115">Office Online</span></span></td>
    <td> <span data-ttu-id="98fee-116">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-116">- Taskpane</span></span><br><span data-ttu-id="98fee-117">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="98fee-117">
        - Content</span></span><br><span data-ttu-id="98fee-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="98fee-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="98fee-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="98fee-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="98fee-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="98fee-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="98fee-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="98fee-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="98fee-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="98fee-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="98fee-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="98fee-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="98fee-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="98fee-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="98fee-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-127">
        -BindingEvents</span></span><br><span data-ttu-id="98fee-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-128">
        -CompressedFile</span></span><br><span data-ttu-id="98fee-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-129">
        -DocumentEvents</span></span><br><span data-ttu-id="98fee-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="98fee-130">
        - File</span></span><br><span data-ttu-id="98fee-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-131">
        -MatrixBindings</span></span><br><span data-ttu-id="98fee-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-132">
        -MatrixCoercion</span></span><br><span data-ttu-id="98fee-133">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="98fee-133">
        - Selection</span></span><br><span data-ttu-id="98fee-134">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-134">
        - Settings</span></span><br><span data-ttu-id="98fee-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-135">
        -TableBindings</span></span><br><span data-ttu-id="98fee-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-136">
        -TableCoercion</span></span><br><span data-ttu-id="98fee-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-137">
        -TextBindings</span></span><br><span data-ttu-id="98fee-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-138">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-139">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="98fee-139">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="98fee-140">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-140">
        - Taskpane</span></span><br><span data-ttu-id="98fee-141">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="98fee-141">
        - Content</span></span></td>
    <td>  <span data-ttu-id="98fee-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="98fee-143">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-143">
        -BindingEvents</span></span><br><span data-ttu-id="98fee-144">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-144">
        -CompressedFile</span></span><br><span data-ttu-id="98fee-145">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-145">
        -DocumentEvents</span></span><br><span data-ttu-id="98fee-146">
        - File</span><span class="sxs-lookup"><span data-stu-id="98fee-146">
        - File</span></span><br><span data-ttu-id="98fee-147">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-147">
        -ImageCoercion</span></span><br><span data-ttu-id="98fee-148">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-148">
        -MatrixBindings</span></span><br><span data-ttu-id="98fee-149">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-149">
        -MatrixCoercion</span></span><br><span data-ttu-id="98fee-150">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="98fee-150">
        - Selection</span></span><br><span data-ttu-id="98fee-151">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-151">
        - Settings</span></span><br><span data-ttu-id="98fee-152">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-152">
        -TableBindings</span></span><br><span data-ttu-id="98fee-153">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-153">
        -TableCoercion</span></span><br><span data-ttu-id="98fee-154">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-154">
        -TextBindings</span></span><br><span data-ttu-id="98fee-155">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-155">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-156">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="98fee-156">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="98fee-157">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-157">- Taskpane</span></span><br><span data-ttu-id="98fee-158">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="98fee-158">
        - Content</span></span><br><span data-ttu-id="98fee-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="98fee-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="98fee-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="98fee-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="98fee-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="98fee-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="98fee-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="98fee-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="98fee-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="98fee-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="98fee-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="98fee-166">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="98fee-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="98fee-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-168">-BindingEvents</span></span><br><span data-ttu-id="98fee-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-169">
        -CompressedFile</span></span><br><span data-ttu-id="98fee-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-170">
        -DocumentEvents</span></span><br><span data-ttu-id="98fee-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="98fee-171">
        - File</span></span><br><span data-ttu-id="98fee-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-172">
        -ImageCoercion</span></span><br><span data-ttu-id="98fee-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-173">
        -MatrixBindings</span></span><br><span data-ttu-id="98fee-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="98fee-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="98fee-175">
        - Selection</span></span><br><span data-ttu-id="98fee-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-176">
        - Settings</span></span><br><span data-ttu-id="98fee-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-177">
        -TableBindings</span></span><br><span data-ttu-id="98fee-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-178">
        -TableCoercion</span></span><br><span data-ttu-id="98fee-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-179">
        -TextBindings</span></span><br><span data-ttu-id="98fee-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-181">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="98fee-181">Office for Windows</span></span></td>
    <td><span data-ttu-id="98fee-182">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-182">- Taskpane</span></span><br><span data-ttu-id="98fee-183">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="98fee-183">
        - Content</span></span><br><span data-ttu-id="98fee-184">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-184">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="98fee-185">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-185">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="98fee-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="98fee-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="98fee-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="98fee-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="98fee-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="98fee-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="98fee-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="98fee-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="98fee-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="98fee-191">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="98fee-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="98fee-193">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-193">-BindingEvents</span></span><br><span data-ttu-id="98fee-194">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-194">
        -CompressedFile</span></span><br><span data-ttu-id="98fee-195">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-195">
        -DocumentEvents</span></span><br><span data-ttu-id="98fee-196">
        - File</span><span class="sxs-lookup"><span data-stu-id="98fee-196">
        - File</span></span><br><span data-ttu-id="98fee-197">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-197">
        -ImageCoercion</span></span><br><span data-ttu-id="98fee-198">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-198">
        -MatrixBindings</span></span><br><span data-ttu-id="98fee-199">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-199">
        -MatrixCoercion</span></span><br><span data-ttu-id="98fee-200">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="98fee-200">
        - Selection</span></span><br><span data-ttu-id="98fee-201">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-201">
        - Settings</span></span><br><span data-ttu-id="98fee-202">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-202">
        -TableBindings</span></span><br><span data-ttu-id="98fee-203">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-203">
        -TableCoercion</span></span><br><span data-ttu-id="98fee-204">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-204">
        -TextBindings</span></span><br><span data-ttu-id="98fee-205">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-205">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-206">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="98fee-206">Office for iOS</span></span></td>
    <td><span data-ttu-id="98fee-207">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-207">- Taskpane</span></span><br><span data-ttu-id="98fee-208">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="98fee-208">
        - Content</span></span></td>
    <td><span data-ttu-id="98fee-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="98fee-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="98fee-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="98fee-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="98fee-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="98fee-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="98fee-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="98fee-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="98fee-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="98fee-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="98fee-215">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="98fee-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="98fee-217">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-217">-BindingEvents</span></span><br><span data-ttu-id="98fee-218">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-218">
        -CompressedFile</span></span><br><span data-ttu-id="98fee-219">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-219">
        -DocumentEvents</span></span><br><span data-ttu-id="98fee-220">
        - File</span><span class="sxs-lookup"><span data-stu-id="98fee-220">
        - File</span></span><br><span data-ttu-id="98fee-221">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-221">
        -ImageCoercion</span></span><br><span data-ttu-id="98fee-222">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-222">
        -MatrixBindings</span></span><br><span data-ttu-id="98fee-223">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-223">
        -MatrixCoercion</span></span><br><span data-ttu-id="98fee-224">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="98fee-224">
        - Selection</span></span><br><span data-ttu-id="98fee-225">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-225">
        - Settings</span></span><br><span data-ttu-id="98fee-226">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-226">
        -TableBindings</span></span><br><span data-ttu-id="98fee-227">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-227">
        -TableCoercion</span></span><br><span data-ttu-id="98fee-228">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-228">
        -TextBindings</span></span><br><span data-ttu-id="98fee-229">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-229">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-230">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="98fee-230">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="98fee-231">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-231">- Taskpane</span></span><br><span data-ttu-id="98fee-232">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="98fee-232">
        - Content</span></span><br><span data-ttu-id="98fee-233">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-233">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="98fee-234">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-234">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="98fee-235">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-235">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="98fee-236">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-236">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="98fee-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="98fee-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="98fee-238">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="98fee-238">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="98fee-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="98fee-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="98fee-240">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="98fee-240">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="98fee-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="98fee-242">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-242">-BindingEvents</span></span><br><span data-ttu-id="98fee-243">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-243">
        -CompressedFile</span></span><br><span data-ttu-id="98fee-244">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-244">
        -DocumentEvents</span></span><br><span data-ttu-id="98fee-245">
        - File</span><span class="sxs-lookup"><span data-stu-id="98fee-245">
        - File</span></span><br><span data-ttu-id="98fee-246">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-246">
        -ImageCoercion</span></span><br><span data-ttu-id="98fee-247">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-247">
        -MatrixBindings</span></span><br><span data-ttu-id="98fee-248">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-248">
        -MatrixCoercion</span></span><br><span data-ttu-id="98fee-249">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="98fee-249">
        -PdfFile</span></span><br><span data-ttu-id="98fee-250">
        - Sélection</span><span class="sxs-lookup"><span data-stu-id="98fee-250">
        - Selection</span></span><br><span data-ttu-id="98fee-251">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-251">
        - Settings</span></span><br><span data-ttu-id="98fee-252">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-252">
        -TableBindings</span></span><br><span data-ttu-id="98fee-253">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-253">
        -TableCoercion</span></span><br><span data-ttu-id="98fee-254">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-254">
        -TextBindings</span></span><br><span data-ttu-id="98fee-255">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-255">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-256">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="98fee-256">Office for Mac</span></span></td>
    <td><span data-ttu-id="98fee-257">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-257">- Taskpane</span></span><br><span data-ttu-id="98fee-258">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="98fee-258">
        - Content</span></span><br><span data-ttu-id="98fee-259">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-259">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="98fee-260">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-260">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="98fee-261">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-261">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="98fee-262">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-262">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="98fee-263">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="98fee-263">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="98fee-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="98fee-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="98fee-265">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="98fee-265">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="98fee-266">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="98fee-266">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="98fee-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="98fee-268">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-268">-BindingEvents</span></span><br><span data-ttu-id="98fee-269">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-269">
        -CompressedFile</span></span><br><span data-ttu-id="98fee-270">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-270">
        -DocumentEvents</span></span><br><span data-ttu-id="98fee-271">
        - File</span><span class="sxs-lookup"><span data-stu-id="98fee-271">
        - File</span></span><br><span data-ttu-id="98fee-272">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-272">
        -ImageCoercion</span></span><br><span data-ttu-id="98fee-273">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-273">
        -MatrixBindings</span></span><br><span data-ttu-id="98fee-274">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-274">
        -MatrixCoercion</span></span><br><span data-ttu-id="98fee-275">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="98fee-275">
        -PdfFile</span></span><br><span data-ttu-id="98fee-276">
        - Sélection</span><span class="sxs-lookup"><span data-stu-id="98fee-276">
        - Selection</span></span><br><span data-ttu-id="98fee-277">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-277">
        - Settings</span></span><br><span data-ttu-id="98fee-278">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-278">
        -TableBindings</span></span><br><span data-ttu-id="98fee-279">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-279">
        -TableCoercion</span></span><br><span data-ttu-id="98fee-280">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-280">
        -TextBindings</span></span><br><span data-ttu-id="98fee-281">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-281">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="98fee-282">Outlook</span><span class="sxs-lookup"><span data-stu-id="98fee-282">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="98fee-283">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="98fee-283">Platform</span></span></th>
    <th><span data-ttu-id="98fee-284">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="98fee-284">Extension points</span></span></th>
    <th><span data-ttu-id="98fee-285">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="98fee-285">API requirement sets</span></span></th>
    <th><span data-ttu-id="98fee-286"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="98fee-286"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-287">Office Online</span><span class="sxs-lookup"><span data-stu-id="98fee-287">Office Online</span></span></td>
    <td> <span data-ttu-id="98fee-288">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="98fee-288">- Mail Read</span></span><br><span data-ttu-id="98fee-289">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="98fee-289">
      - Mail Compose</span></span><br><span data-ttu-id="98fee-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="98fee-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="98fee-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="98fee-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="98fee-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="98fee-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="98fee-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="98fee-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="98fee-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="98fee-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="98fee-297">Non disponible</span><span class="sxs-lookup"><span data-stu-id="98fee-297">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-298">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="98fee-298">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="98fee-299">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="98fee-299">- Mail Read</span></span><br><span data-ttu-id="98fee-300">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="98fee-300">
      - Mail Compose</span></span><br><span data-ttu-id="98fee-301">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-301">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="98fee-302">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-302">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="98fee-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="98fee-304">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-304">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="98fee-305">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="98fee-305">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="98fee-306">Non disponible</span><span class="sxs-lookup"><span data-stu-id="98fee-306">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-307">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="98fee-307">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="98fee-308">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="98fee-308">- Mail Read</span></span><br><span data-ttu-id="98fee-309">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="98fee-309">
      - Mail Compose</span></span><br><span data-ttu-id="98fee-310">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-310">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="98fee-311">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="98fee-311">
      - Modules</span></span></td>
    <td> <span data-ttu-id="98fee-312">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-312">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="98fee-313">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-313">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="98fee-314">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-314">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="98fee-315">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="98fee-315">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="98fee-316">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="98fee-316">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="98fee-317">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="98fee-317">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="98fee-318">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="98fee-318">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="98fee-319">Non disponible</span><span class="sxs-lookup"><span data-stu-id="98fee-319">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-320">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="98fee-320">Office for Windows</span></span></td>
    <td> <span data-ttu-id="98fee-321">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="98fee-321">- Mail Read</span></span><br><span data-ttu-id="98fee-322">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="98fee-322">
      - Mail Compose</span></span><br><span data-ttu-id="98fee-323">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-323">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="98fee-324">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="98fee-324">
      - Modules</span></span></td>
    <td> <span data-ttu-id="98fee-325">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-325">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="98fee-326">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-326">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="98fee-327">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-327">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="98fee-328">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="98fee-328">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="98fee-329">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="98fee-329">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="98fee-330">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="98fee-330">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="98fee-331">Non disponible</span><span class="sxs-lookup"><span data-stu-id="98fee-331">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-332">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="98fee-332">Office for iOS</span></span></td>
    <td> <span data-ttu-id="98fee-333">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="98fee-333">- Mail Read</span></span><br><span data-ttu-id="98fee-334">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-334">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="98fee-335">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-335">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="98fee-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="98fee-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="98fee-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="98fee-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="98fee-339">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="98fee-339">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="98fee-340">Non disponible</span><span class="sxs-lookup"><span data-stu-id="98fee-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-341">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="98fee-341">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="98fee-342">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="98fee-342">- Mail Read</span></span><br><span data-ttu-id="98fee-343">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="98fee-343">
      - Mail Compose</span></span><br><span data-ttu-id="98fee-344">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-344">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="98fee-345">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-345">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="98fee-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="98fee-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="98fee-348">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="98fee-348">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="98fee-349">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="98fee-349">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="98fee-350">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="98fee-350">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="98fee-351">Non disponible</span><span class="sxs-lookup"><span data-stu-id="98fee-351">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-352">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="98fee-352">Office for Mac</span></span></td>
    <td> <span data-ttu-id="98fee-353">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="98fee-353">- Mail Read</span></span><br><span data-ttu-id="98fee-354">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="98fee-354">
      - Mail Compose</span></span><br><span data-ttu-id="98fee-355">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-355">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="98fee-356">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-356">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="98fee-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="98fee-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="98fee-359">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="98fee-359">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="98fee-360">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="98fee-360">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="98fee-361">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="98fee-361">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="98fee-362">Non disponible</span><span class="sxs-lookup"><span data-stu-id="98fee-362">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-363">Office pour Android</span><span class="sxs-lookup"><span data-stu-id="98fee-363">Office for Android</span></span></td>
    <td> <span data-ttu-id="98fee-364">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="98fee-364">- Mail Read</span></span><br><span data-ttu-id="98fee-365">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-365">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="98fee-366">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-366">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="98fee-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="98fee-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="98fee-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="98fee-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="98fee-370">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="98fee-370">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="98fee-371">Non disponible</span><span class="sxs-lookup"><span data-stu-id="98fee-371">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="98fee-372">Word</span><span class="sxs-lookup"><span data-stu-id="98fee-372">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="98fee-373">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="98fee-373">Platform</span></span></th>
    <th><span data-ttu-id="98fee-374">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="98fee-374">Extension points</span></span></th>
    <th><span data-ttu-id="98fee-375">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="98fee-375">API requirement sets</span></span></th>
    <th><span data-ttu-id="98fee-376"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="98fee-376"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-377">Office Online</span><span class="sxs-lookup"><span data-stu-id="98fee-377">Office Online</span></span></td>
    <td> <span data-ttu-id="98fee-378">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-378">- Taskpane</span></span><br><span data-ttu-id="98fee-379">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-379">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="98fee-380">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-380">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="98fee-381">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-381">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="98fee-382">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-382">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="98fee-383">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-383">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="98fee-384">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-384">-BindingEvents</span></span><br><span data-ttu-id="98fee-385">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="98fee-385">
         -</span></span><br><span data-ttu-id="98fee-386">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-386">
         -DocumentEvents</span></span><br><span data-ttu-id="98fee-387">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="98fee-387">
         - File</span></span><br><span data-ttu-id="98fee-388">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-388">
         -HtmlCoercion</span></span><br><span data-ttu-id="98fee-389">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-389">
         -ImageCoercion</span></span><br><span data-ttu-id="98fee-390">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-390">
         -MatrixBindings</span></span><br><span data-ttu-id="98fee-391">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-391">
         -MatrixCoercion</span></span><br><span data-ttu-id="98fee-392">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-392">
         -OoxmlCoercion</span></span><br><span data-ttu-id="98fee-393">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="98fee-393">
         -PdfFile</span></span><br><span data-ttu-id="98fee-394">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="98fee-394">
         - Selection</span></span><br><span data-ttu-id="98fee-395">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-395">
         - Settings</span></span><br><span data-ttu-id="98fee-396">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-396">
         -TableBindings</span></span><br><span data-ttu-id="98fee-397">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-397">
         -TableCoercion</span></span><br><span data-ttu-id="98fee-398">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-398">
         -TextBindings</span></span><br><span data-ttu-id="98fee-399">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-399">
         -TextCoercion</span></span><br><span data-ttu-id="98fee-400">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="98fee-400">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-401">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="98fee-401">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="98fee-402">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-402">- Taskpane</span></span></td>
    <td> <span data-ttu-id="98fee-403">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-403">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="98fee-404">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-404">-BindingEvents</span></span><br><span data-ttu-id="98fee-405">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-405">
         -CompressedFile</span></span><br><span data-ttu-id="98fee-406">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="98fee-406">
         -</span></span><br><span data-ttu-id="98fee-407">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-407">
         -DocumentEvents</span></span><br><span data-ttu-id="98fee-408">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="98fee-408">
         - File</span></span><br><span data-ttu-id="98fee-409">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-409">
         -HtmlCoercion</span></span><br><span data-ttu-id="98fee-410">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-410">
         -ImageCoercion</span></span><br><span data-ttu-id="98fee-411">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-411">
         -MatrixBindings</span></span><br><span data-ttu-id="98fee-412">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-412">
         -MatrixCoercion</span></span><br><span data-ttu-id="98fee-413">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-413">
         -OoxmlCoercion</span></span><br><span data-ttu-id="98fee-414">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="98fee-414">
         -PdfFile</span></span><br><span data-ttu-id="98fee-415">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="98fee-415">
         - Selection</span></span><br><span data-ttu-id="98fee-416">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-416">
         - Settings</span></span><br><span data-ttu-id="98fee-417">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-417">
         -TableBindings</span></span><br><span data-ttu-id="98fee-418">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-418">
         -TableCoercion</span></span><br><span data-ttu-id="98fee-419">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-419">
         -TextBindings</span></span><br><span data-ttu-id="98fee-420">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-420">
         -TextCoercion</span></span><br><span data-ttu-id="98fee-421">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="98fee-421">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-422">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="98fee-422">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="98fee-423">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-423">- Taskpane</span></span><br><span data-ttu-id="98fee-424">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-424">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="98fee-425">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-425">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="98fee-426">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-426">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="98fee-427">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-427">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="98fee-428">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-428">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="98fee-429">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-429">-BindingEvents</span></span><br><span data-ttu-id="98fee-430">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-430">
         -CompressedFile</span></span><br><span data-ttu-id="98fee-431">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="98fee-431">
         -</span></span><br><span data-ttu-id="98fee-432">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-432">
         -DocumentEvents</span></span><br><span data-ttu-id="98fee-433">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="98fee-433">
         - File</span></span><br><span data-ttu-id="98fee-434">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-434">
         -HtmlCoercion</span></span><br><span data-ttu-id="98fee-435">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-435">
         -ImageCoercion</span></span><br><span data-ttu-id="98fee-436">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-436">
         -MatrixBindings</span></span><br><span data-ttu-id="98fee-437">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-437">
         -MatrixCoercion</span></span><br><span data-ttu-id="98fee-438">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-438">
         -OoxmlCoercion</span></span><br><span data-ttu-id="98fee-439">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="98fee-439">
         -PdfFile</span></span><br><span data-ttu-id="98fee-440">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="98fee-440">
         - Selection</span></span><br><span data-ttu-id="98fee-441">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-441">
         - Settings</span></span><br><span data-ttu-id="98fee-442">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-442">
         -TableBindings</span></span><br><span data-ttu-id="98fee-443">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-443">
         -TableCoercion</span></span><br><span data-ttu-id="98fee-444">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-444">
         -TextBindings</span></span><br><span data-ttu-id="98fee-445">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-445">
         -TextCoercion</span></span><br><span data-ttu-id="98fee-446">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="98fee-446">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-447">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="98fee-447">Office for Windows</span></span></td>
    <td> <span data-ttu-id="98fee-448">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-448">- Taskpane</span></span><br><span data-ttu-id="98fee-449">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-449">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="98fee-450">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-450">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="98fee-451">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-451">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="98fee-452">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-452">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="98fee-453">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-453">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="98fee-454">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-454">-BindingEvents</span></span><br><span data-ttu-id="98fee-455">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-455">
         -CompressedFile</span></span><br><span data-ttu-id="98fee-456">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="98fee-456">
         -</span></span><br><span data-ttu-id="98fee-457">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-457">
         -DocumentEvents</span></span><br><span data-ttu-id="98fee-458">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="98fee-458">
         - File</span></span><br><span data-ttu-id="98fee-459">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-459">
         -HtmlCoercion</span></span><br><span data-ttu-id="98fee-460">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-460">
         -ImageCoercion</span></span><br><span data-ttu-id="98fee-461">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-461">
         -MatrixBindings</span></span><br><span data-ttu-id="98fee-462">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-462">
         -MatrixCoercion</span></span><br><span data-ttu-id="98fee-463">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-463">
         -OoxmlCoercion</span></span><br><span data-ttu-id="98fee-464">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="98fee-464">
         -PdfFile</span></span><br><span data-ttu-id="98fee-465">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="98fee-465">
         - Selection</span></span><br><span data-ttu-id="98fee-466">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-466">
         - Settings</span></span><br><span data-ttu-id="98fee-467">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-467">
         -TableBindings</span></span><br><span data-ttu-id="98fee-468">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-468">
         -TableCoercion</span></span><br><span data-ttu-id="98fee-469">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-469">
         -TextBindings</span></span><br><span data-ttu-id="98fee-470">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-470">
         -TextCoercion</span></span><br><span data-ttu-id="98fee-471">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="98fee-471">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-472">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="98fee-472">Office for iOS</span></span></td>
    <td> <span data-ttu-id="98fee-473">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-473">- Taskpane</span></span></td>
    <td> <span data-ttu-id="98fee-474">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-474">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="98fee-475">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-475">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="98fee-476">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-476">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="98fee-477">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="98fee-477">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="98fee-478">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-478">-BindingEvents</span></span><br><span data-ttu-id="98fee-479">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-479">
         -CompressedFile</span></span><br><span data-ttu-id="98fee-480">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="98fee-480">
         -</span></span><br><span data-ttu-id="98fee-481">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-481">
         -DocumentEvents</span></span><br><span data-ttu-id="98fee-482">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="98fee-482">
         - File</span></span><br><span data-ttu-id="98fee-483">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-483">
         -HtmlCoercion</span></span><br><span data-ttu-id="98fee-484">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-484">
         -ImageCoercion</span></span><br><span data-ttu-id="98fee-485">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-485">
         -MatrixBindings</span></span><br><span data-ttu-id="98fee-486">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-486">
         -MatrixCoercion</span></span><br><span data-ttu-id="98fee-487">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-487">
         -OoxmlCoercion</span></span><br><span data-ttu-id="98fee-488">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="98fee-488">
         -PdfFile</span></span><br><span data-ttu-id="98fee-489">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="98fee-489">
         - Selection</span></span><br><span data-ttu-id="98fee-490">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-490">
         - Settings</span></span><br><span data-ttu-id="98fee-491">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-491">
         -TableBindings</span></span><br><span data-ttu-id="98fee-492">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-492">
         -TableCoercion</span></span><br><span data-ttu-id="98fee-493">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-493">
         -TextBindings</span></span><br><span data-ttu-id="98fee-494">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-494">
         -TextCoercion</span></span><br><span data-ttu-id="98fee-495">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="98fee-495">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-496">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="98fee-496">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="98fee-497">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-497">- Taskpane</span></span><br><span data-ttu-id="98fee-498">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-498">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="98fee-499">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-499">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="98fee-500">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-500">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="98fee-501">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-501">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="98fee-502">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="98fee-502">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="98fee-503">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-503">-BindingEvents</span></span><br><span data-ttu-id="98fee-504">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-504">
         -CompressedFile</span></span><br><span data-ttu-id="98fee-505">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="98fee-505">
         -</span></span><br><span data-ttu-id="98fee-506">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-506">
         -DocumentEvents</span></span><br><span data-ttu-id="98fee-507">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="98fee-507">
         - File</span></span><br><span data-ttu-id="98fee-508">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-508">
         -HtmlCoercion</span></span><br><span data-ttu-id="98fee-509">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-509">
         -ImageCoercion</span></span><br><span data-ttu-id="98fee-510">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-510">
         -MatrixBindings</span></span><br><span data-ttu-id="98fee-511">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-511">
         -MatrixCoercion</span></span><br><span data-ttu-id="98fee-512">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-512">
         -OoxmlCoercion</span></span><br><span data-ttu-id="98fee-513">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="98fee-513">
         -PdfFile</span></span><br><span data-ttu-id="98fee-514">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="98fee-514">
         - Selection</span></span><br><span data-ttu-id="98fee-515">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-515">
         - Settings</span></span><br><span data-ttu-id="98fee-516">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-516">
         -TableBindings</span></span><br><span data-ttu-id="98fee-517">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-517">
         -TableCoercion</span></span><br><span data-ttu-id="98fee-518">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-518">
         -TextBindings</span></span><br><span data-ttu-id="98fee-519">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-519">
         -TextCoercion</span></span><br><span data-ttu-id="98fee-520">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="98fee-520">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-521">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="98fee-521">Office for Mac</span></span></td>
    <td> <span data-ttu-id="98fee-522">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-522">- Taskpane</span></span><br><span data-ttu-id="98fee-523">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-523">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="98fee-524">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-524">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="98fee-525">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="98fee-525">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="98fee-526">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="98fee-526">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="98fee-527">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="98fee-527">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="98fee-528">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-528">-BindingEvents</span></span><br><span data-ttu-id="98fee-529">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-529">
         -CompressedFile</span></span><br><span data-ttu-id="98fee-530">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="98fee-530">
         -</span></span><br><span data-ttu-id="98fee-531">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-531">
         -DocumentEvents</span></span><br><span data-ttu-id="98fee-532">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="98fee-532">
         - File</span></span><br><span data-ttu-id="98fee-533">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-533">
         -HtmlCoercion</span></span><br><span data-ttu-id="98fee-534">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-534">
         -ImageCoercion</span></span><br><span data-ttu-id="98fee-535">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-535">
         -MatrixBindings</span></span><br><span data-ttu-id="98fee-536">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-536">
         -MatrixCoercion</span></span><br><span data-ttu-id="98fee-537">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-537">
         -OoxmlCoercion</span></span><br><span data-ttu-id="98fee-538">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="98fee-538">
         -PdfFile</span></span><br><span data-ttu-id="98fee-539">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="98fee-539">
         - Selection</span></span><br><span data-ttu-id="98fee-540">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-540">
         - Settings</span></span><br><span data-ttu-id="98fee-541">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-541">
         -TableBindings</span></span><br><span data-ttu-id="98fee-542">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-542">
         -TableCoercion</span></span><br><span data-ttu-id="98fee-543">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="98fee-543">
         -TextBindings</span></span><br><span data-ttu-id="98fee-544">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-544">
         -TextCoercion</span></span><br><span data-ttu-id="98fee-545">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="98fee-545">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="98fee-546">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="98fee-546">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="98fee-547">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="98fee-547">Platform</span></span></th>
    <th><span data-ttu-id="98fee-548">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="98fee-548">Extension points</span></span></th>
    <th><span data-ttu-id="98fee-549">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="98fee-549">API requirement sets</span></span></th>
    <th><span data-ttu-id="98fee-550"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="98fee-550"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-551">Office Online</span><span class="sxs-lookup"><span data-stu-id="98fee-551">Office Online</span></span></td>
    <td> <span data-ttu-id="98fee-552">- Contenu</span><span class="sxs-lookup"><span data-stu-id="98fee-552">- Content</span></span><br><span data-ttu-id="98fee-553">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-553">
         - Taskpane</span></span><br><span data-ttu-id="98fee-554">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-554">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="98fee-555">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-555">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="98fee-556">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="98fee-556">-ActiveView</span></span><br><span data-ttu-id="98fee-557">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-557">
         -CompressedFile</span></span><br><span data-ttu-id="98fee-558">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-558">
         -DocumentEvents</span></span><br><span data-ttu-id="98fee-559">
         - File</span><span class="sxs-lookup"><span data-stu-id="98fee-559">
         - File</span></span><br><span data-ttu-id="98fee-560">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-560">
         -ImageCoercion</span></span><br><span data-ttu-id="98fee-561">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="98fee-561">
         -PdfFile</span></span><br><span data-ttu-id="98fee-562">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="98fee-562">
         - Selection</span></span><br><span data-ttu-id="98fee-563">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-563">
         - Settings</span></span><br><span data-ttu-id="98fee-564">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-564">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-565">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="98fee-565">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="98fee-566">- Contenu</span><span class="sxs-lookup"><span data-stu-id="98fee-566">- Content</span></span><br><span data-ttu-id="98fee-567">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-567">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="98fee-568">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="98fee-568">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="98fee-569">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="98fee-569">-ActiveView</span></span><br><span data-ttu-id="98fee-570">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-570">
         -CompressedFile</span></span><br><span data-ttu-id="98fee-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-571">
         -DocumentEvents</span></span><br><span data-ttu-id="98fee-572">
         - File</span><span class="sxs-lookup"><span data-stu-id="98fee-572">
         - File</span></span><br><span data-ttu-id="98fee-573">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-573">
         -ImageCoercion</span></span><br><span data-ttu-id="98fee-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="98fee-574">
         -PdfFile</span></span><br><span data-ttu-id="98fee-575">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="98fee-575">
         - Selection</span></span><br><span data-ttu-id="98fee-576">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-576">
         - Settings</span></span><br><span data-ttu-id="98fee-577">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-577">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-578">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="98fee-578">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="98fee-579">- Contenu</span><span class="sxs-lookup"><span data-stu-id="98fee-579">- Content</span></span><br><span data-ttu-id="98fee-580">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-580">
         - Taskpane</span></span><br><span data-ttu-id="98fee-581">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-581">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="98fee-582">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-582">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="98fee-583">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="98fee-583">-ActiveView</span></span><br><span data-ttu-id="98fee-584">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-584">
         -CompressedFile</span></span><br><span data-ttu-id="98fee-585">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-585">
         -DocumentEvents</span></span><br><span data-ttu-id="98fee-586">
         - File</span><span class="sxs-lookup"><span data-stu-id="98fee-586">
         - File</span></span><br><span data-ttu-id="98fee-587">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-587">
         -ImageCoercion</span></span><br><span data-ttu-id="98fee-588">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="98fee-588">
         -PdfFile</span></span><br><span data-ttu-id="98fee-589">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="98fee-589">
         - Selection</span></span><br><span data-ttu-id="98fee-590">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-590">
         - Settings</span></span><br><span data-ttu-id="98fee-591">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-591">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-592">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="98fee-592">Office for Windows</span></span></td>
    <td> <span data-ttu-id="98fee-593">- Contenu</span><span class="sxs-lookup"><span data-stu-id="98fee-593">- Content</span></span><br><span data-ttu-id="98fee-594">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-594">
         - Taskpane</span></span><br><span data-ttu-id="98fee-595">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-595">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="98fee-596">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-596">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="98fee-597">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="98fee-597">-ActiveView</span></span><br><span data-ttu-id="98fee-598">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-598">
         -CompressedFile</span></span><br><span data-ttu-id="98fee-599">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-599">
         -DocumentEvents</span></span><br><span data-ttu-id="98fee-600">
         - File</span><span class="sxs-lookup"><span data-stu-id="98fee-600">
         - File</span></span><br><span data-ttu-id="98fee-601">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-601">
         -ImageCoercion</span></span><br><span data-ttu-id="98fee-602">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="98fee-602">
         -PdfFile</span></span><br><span data-ttu-id="98fee-603">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="98fee-603">
         - Selection</span></span><br><span data-ttu-id="98fee-604">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-604">
         - Settings</span></span><br><span data-ttu-id="98fee-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-605">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-606">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="98fee-606">Office for iOS</span></span></td>
    <td> <span data-ttu-id="98fee-607">- Contenu</span><span class="sxs-lookup"><span data-stu-id="98fee-607">- Content</span></span><br><span data-ttu-id="98fee-608">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-608">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="98fee-609">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-609">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="98fee-610">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="98fee-610">-ActiveView</span></span><br><span data-ttu-id="98fee-611">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-611">
         -CompressedFile</span></span><br><span data-ttu-id="98fee-612">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-612">
         -DocumentEvents</span></span><br><span data-ttu-id="98fee-613">
         - File</span><span class="sxs-lookup"><span data-stu-id="98fee-613">
         - File</span></span><br><span data-ttu-id="98fee-614">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="98fee-614">
         -PdfFile</span></span><br><span data-ttu-id="98fee-615">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="98fee-615">
         - Selection</span></span><br><span data-ttu-id="98fee-616">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-616">
         - Settings</span></span><br><span data-ttu-id="98fee-617">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-617">
         -TextCoercion</span></span><br><span data-ttu-id="98fee-618">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-618">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-619">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="98fee-619">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="98fee-620">- Contenu</span><span class="sxs-lookup"><span data-stu-id="98fee-620">- Content</span></span><br><span data-ttu-id="98fee-621">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-621">
         - Taskpane</span></span><br><span data-ttu-id="98fee-622">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-622">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="98fee-623">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-623">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="98fee-624">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="98fee-624">-ActiveView</span></span><br><span data-ttu-id="98fee-625">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-625">
         -CompressedFile</span></span><br><span data-ttu-id="98fee-626">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-626">
         -DocumentEvents</span></span><br><span data-ttu-id="98fee-627">
         - File</span><span class="sxs-lookup"><span data-stu-id="98fee-627">
         - File</span></span><br><span data-ttu-id="98fee-628">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-628">
         -ImageCoercion</span></span><br><span data-ttu-id="98fee-629">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="98fee-629">
         -PdfFile</span></span><br><span data-ttu-id="98fee-630">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="98fee-630">
         - Selection</span></span><br><span data-ttu-id="98fee-631">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-631">
         - Settings</span></span><br><span data-ttu-id="98fee-632">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-632">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-633">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="98fee-633">Office for Mac</span></span></td>
    <td> <span data-ttu-id="98fee-634">- Contenu</span><span class="sxs-lookup"><span data-stu-id="98fee-634">- Content</span></span><br><span data-ttu-id="98fee-635">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-635">
         - Taskpane</span></span><br><span data-ttu-id="98fee-636">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-636">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="98fee-637">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-637">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="98fee-638">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="98fee-638">-ActiveView</span></span><br><span data-ttu-id="98fee-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="98fee-639">
         -CompressedFile</span></span><br><span data-ttu-id="98fee-640">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-640">
         -DocumentEvents</span></span><br><span data-ttu-id="98fee-641">
         - File</span><span class="sxs-lookup"><span data-stu-id="98fee-641">
         - File</span></span><br><span data-ttu-id="98fee-642">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-642">
         -ImageCoercion</span></span><br><span data-ttu-id="98fee-643">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="98fee-643">
         -PdfFile</span></span><br><span data-ttu-id="98fee-644">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="98fee-644">
         - Selection</span></span><br><span data-ttu-id="98fee-645">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-645">
         - Settings</span></span><br><span data-ttu-id="98fee-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-646">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="98fee-647">OneNote</span><span class="sxs-lookup"><span data-stu-id="98fee-647">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="98fee-648">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="98fee-648">Platform</span></span></th>
    <th><span data-ttu-id="98fee-649">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="98fee-649">Extension points</span></span></th>
    <th><span data-ttu-id="98fee-650">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="98fee-650">API requirement sets</span></span></th>
    <th><span data-ttu-id="98fee-651"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="98fee-651"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="98fee-652">Office Online</span><span class="sxs-lookup"><span data-stu-id="98fee-652">Office Online</span></span></td>
    <td> <span data-ttu-id="98fee-653">- Contenu</span><span class="sxs-lookup"><span data-stu-id="98fee-653">- Content</span></span><br><span data-ttu-id="98fee-654">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="98fee-654">
         - Taskpane</span></span><br><span data-ttu-id="98fee-655">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="98fee-655">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="98fee-656">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-656">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="98fee-657">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="98fee-657">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="98fee-658">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="98fee-658">-DocumentEvents</span></span><br><span data-ttu-id="98fee-659">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-659">
         -HtmlCoercion</span></span><br><span data-ttu-id="98fee-660">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-660">
         -ImageCoercion</span></span><br><span data-ttu-id="98fee-661">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="98fee-661">
         - Settings</span></span><br><span data-ttu-id="98fee-662">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="98fee-662">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="98fee-663">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="98fee-663">See also</span></span>

- [<span data-ttu-id="98fee-664">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="98fee-664">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="98fee-665">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="98fee-665">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="98fee-666">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="98fee-666">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="98fee-667">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="98fee-667">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)
