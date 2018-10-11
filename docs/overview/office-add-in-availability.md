---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint et OneNote.
ms.date: 10/03/2018
ms.openlocfilehash: 6f7b5b565773457e6cd8a9eee69eb304784a29a9
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459314"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="52a0b-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="52a0b-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="52a0b-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d’un hôte Office spécifique, d’un ensemble de conditions requises, d’un membre d’API ou d’une version de l’API. Les tableaux suivants contiennent la plateforme disponible, les points d’extension, les ensembles d’API requis et les ensembles d’API courantes requis qui sont actuellement pris en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="52a0b-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="52a0b-p102">Si une cellule de tableau contient un astérisque (\*), cela signifie que nous travaillons dessus. Pour les ensembles de conditions requises pour Projet ou Access, voir [Ensembles de conditions requises communs à Office](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="52a0b-p102">If a table cell contains an asterisk ( \* ), that means we’re working on it. For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="52a0b-p103">Le numéro de build pour Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que les ensembles de conditions requises des API communes, ExcelApi 1.1 et WordApi 1.1.</span><span class="sxs-lookup"><span data-stu-id="52a0b-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="52a0b-110">Excel</span><span class="sxs-lookup"><span data-stu-id="52a0b-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="52a0b-111">Plateforme</span><span class="sxs-lookup"><span data-stu-id="52a0b-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="52a0b-112">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="52a0b-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="52a0b-113">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="52a0b-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="52a0b-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="52a0b-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="52a0b-115">Office Online</span></span></td>
    <td> <span data-ttu-id="52a0b-116">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-116">- Taskpane</span></span><br><span data-ttu-id="52a0b-117">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="52a0b-117">
        - Content</span></span><br><span data-ttu-id="52a0b-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="52a0b-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="52a0b-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="52a0b-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="52a0b-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="52a0b-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="52a0b-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="52a0b-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="52a0b-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="52a0b-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="52a0b-127">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-127">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="52a0b-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-128">
        -BindingEvents</span></span><br><span data-ttu-id="52a0b-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-129">
        -CompressedFile</span></span><br><span data-ttu-id="52a0b-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-130">
        -DocumentEvents</span></span><br><span data-ttu-id="52a0b-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="52a0b-131">
        - File</span></span><br><span data-ttu-id="52a0b-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-132">
        -MatrixBindings</span></span><br><span data-ttu-id="52a0b-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-133">
        -MatrixCoercion</span></span><br><span data-ttu-id="52a0b-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="52a0b-134">
        - Selection</span></span><br><span data-ttu-id="52a0b-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-135">
        - Settings</span></span><br><span data-ttu-id="52a0b-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-136">
        -TableBindings</span></span><br><span data-ttu-id="52a0b-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-137">
        -TableCoercion</span></span><br><span data-ttu-id="52a0b-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-138">
        -TextBindings</span></span><br><span data-ttu-id="52a0b-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-139">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-140">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="52a0b-140">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="52a0b-141">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-141">
        - Taskpane</span></span><br><span data-ttu-id="52a0b-142">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="52a0b-142">
        - Content</span></span></td>
    <td>  <span data-ttu-id="52a0b-143">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-143">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="52a0b-144">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-144">
        -BindingEvents</span></span><br><span data-ttu-id="52a0b-145">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-145">
        -CompressedFile</span></span><br><span data-ttu-id="52a0b-146">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-146">
        -DocumentEvents</span></span><br><span data-ttu-id="52a0b-147">
        - File</span><span class="sxs-lookup"><span data-stu-id="52a0b-147">
        - File</span></span><br><span data-ttu-id="52a0b-148">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-148">
        -ImageCoercion</span></span><br><span data-ttu-id="52a0b-149">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-149">
        -MatrixBindings</span></span><br><span data-ttu-id="52a0b-150">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-150">
        -MatrixCoercion</span></span><br><span data-ttu-id="52a0b-151">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="52a0b-151">
        - Selection</span></span><br><span data-ttu-id="52a0b-152">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-152">
        - Settings</span></span><br><span data-ttu-id="52a0b-153">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-153">
        -TableBindings</span></span><br><span data-ttu-id="52a0b-154">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-154">
        -TableCoercion</span></span><br><span data-ttu-id="52a0b-155">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-155">
        -TextBindings</span></span><br><span data-ttu-id="52a0b-156">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-156">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-157">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="52a0b-157">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="52a0b-158">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-158">- Taskpane</span></span><br><span data-ttu-id="52a0b-159">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="52a0b-159">
        - Content</span></span><br><span data-ttu-id="52a0b-160">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-160">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="52a0b-161">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-161">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="52a0b-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="52a0b-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="52a0b-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="52a0b-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="52a0b-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="52a0b-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-167">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="52a0b-168">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-168">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="52a0b-169">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-169">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="52a0b-170">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-170">-BindingEvents</span></span><br><span data-ttu-id="52a0b-171">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-171">
        -CompressedFile</span></span><br><span data-ttu-id="52a0b-172">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-172">
        -DocumentEvents</span></span><br><span data-ttu-id="52a0b-173">
        - File</span><span class="sxs-lookup"><span data-stu-id="52a0b-173">
        - File</span></span><br><span data-ttu-id="52a0b-174">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-174">
        -ImageCoercion</span></span><br><span data-ttu-id="52a0b-175">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-175">
        -MatrixBindings</span></span><br><span data-ttu-id="52a0b-176">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-176">
        -MatrixCoercion</span></span><br><span data-ttu-id="52a0b-177">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="52a0b-177">
        - Selection</span></span><br><span data-ttu-id="52a0b-178">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-178">
        - Settings</span></span><br><span data-ttu-id="52a0b-179">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-179">
        -TableBindings</span></span><br><span data-ttu-id="52a0b-180">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-180">
        -TableCoercion</span></span><br><span data-ttu-id="52a0b-181">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-181">
        -TextBindings</span></span><br><span data-ttu-id="52a0b-182">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-182">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-183">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="52a0b-183">Office for Windows Desktop.</span></span></td>
    <td><span data-ttu-id="52a0b-184">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-184">- Taskpane</span></span><br><span data-ttu-id="52a0b-185">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="52a0b-185">
        - Content</span></span><br><span data-ttu-id="52a0b-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="52a0b-187">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-187">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="52a0b-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="52a0b-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="52a0b-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="52a0b-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="52a0b-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="52a0b-193">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-193">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="52a0b-194">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-194">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="52a0b-195">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-195">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="52a0b-196">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-196">-BindingEvents</span></span><br><span data-ttu-id="52a0b-197">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-197">
        -CompressedFile</span></span><br><span data-ttu-id="52a0b-198">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-198">
        -DocumentEvents</span></span><br><span data-ttu-id="52a0b-199">
        - File</span><span class="sxs-lookup"><span data-stu-id="52a0b-199">
        - File</span></span><br><span data-ttu-id="52a0b-200">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-200">
        -ImageCoercion</span></span><br><span data-ttu-id="52a0b-201">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-201">
        -MatrixBindings</span></span><br><span data-ttu-id="52a0b-202">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-202">
        -MatrixCoercion</span></span><br><span data-ttu-id="52a0b-203">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="52a0b-203">
        - Selection</span></span><br><span data-ttu-id="52a0b-204">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-204">
        - Settings</span></span><br><span data-ttu-id="52a0b-205">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-205">
        -TableBindings</span></span><br><span data-ttu-id="52a0b-206">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-206">
        -TableCoercion</span></span><br><span data-ttu-id="52a0b-207">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-207">
        -TextBindings</span></span><br><span data-ttu-id="52a0b-208">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-208">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-209">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="52a0b-209">Office for iOS</span></span></td>
    <td><span data-ttu-id="52a0b-210">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-210">- Taskpane</span></span><br><span data-ttu-id="52a0b-211">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="52a0b-211">
        - Content</span></span></td>
    <td><span data-ttu-id="52a0b-212">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-212">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="52a0b-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="52a0b-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="52a0b-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="52a0b-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="52a0b-217">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-217">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="52a0b-218">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-218">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="52a0b-219">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-219">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="52a0b-220">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-220">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="52a0b-221">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-221">-BindingEvents</span></span><br><span data-ttu-id="52a0b-222">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-222">
        -CompressedFile</span></span><br><span data-ttu-id="52a0b-223">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-223">
        -DocumentEvents</span></span><br><span data-ttu-id="52a0b-224">
        - File</span><span class="sxs-lookup"><span data-stu-id="52a0b-224">
        - File</span></span><br><span data-ttu-id="52a0b-225">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-225">
        -ImageCoercion</span></span><br><span data-ttu-id="52a0b-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-226">
        -MatrixBindings</span></span><br><span data-ttu-id="52a0b-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-227">
        -MatrixCoercion</span></span><br><span data-ttu-id="52a0b-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="52a0b-228">
        - Selection</span></span><br><span data-ttu-id="52a0b-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-229">
        - Settings</span></span><br><span data-ttu-id="52a0b-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-230">
        -TableBindings</span></span><br><span data-ttu-id="52a0b-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-231">
        -TableCoercion</span></span><br><span data-ttu-id="52a0b-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-232">
        -TextBindings</span></span><br><span data-ttu-id="52a0b-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-233">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-234">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="52a0b-234">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="52a0b-235">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-235">- Taskpane</span></span><br><span data-ttu-id="52a0b-236">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="52a0b-236">
        - Content</span></span><br><span data-ttu-id="52a0b-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="52a0b-238">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-238">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="52a0b-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="52a0b-240">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-240">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="52a0b-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="52a0b-242">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-242">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="52a0b-243">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-243">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="52a0b-244">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-244">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="52a0b-245">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-245">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="52a0b-246">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-246">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="52a0b-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-247">-BindingEvents</span></span><br><span data-ttu-id="52a0b-248">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-248">
        -CompressedFile</span></span><br><span data-ttu-id="52a0b-249">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-249">
        -DocumentEvents</span></span><br><span data-ttu-id="52a0b-250">
        - File</span><span class="sxs-lookup"><span data-stu-id="52a0b-250">
        - File</span></span><br><span data-ttu-id="52a0b-251">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-251">
        -ImageCoercion</span></span><br><span data-ttu-id="52a0b-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-252">
        -MatrixBindings</span></span><br><span data-ttu-id="52a0b-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-253">
        -MatrixCoercion</span></span><br><span data-ttu-id="52a0b-254">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-254">
        -PdfFile</span></span><br><span data-ttu-id="52a0b-255">
        - Sélection</span><span class="sxs-lookup"><span data-stu-id="52a0b-255">
        - Selection</span></span><br><span data-ttu-id="52a0b-256">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-256">
        - Settings</span></span><br><span data-ttu-id="52a0b-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-257">
        -TableBindings</span></span><br><span data-ttu-id="52a0b-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-258">
        -TableCoercion</span></span><br><span data-ttu-id="52a0b-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-259">
        -TextBindings</span></span><br><span data-ttu-id="52a0b-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-260">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-261">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="52a0b-261">Office for Mac</span></span></td>
    <td><span data-ttu-id="52a0b-262">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-262">- Taskpane</span></span><br><span data-ttu-id="52a0b-263">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="52a0b-263">
        - Content</span></span><br><span data-ttu-id="52a0b-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="52a0b-265">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-265">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="52a0b-266">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-266">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="52a0b-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="52a0b-268">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-268">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="52a0b-269">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-269">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="52a0b-270">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-270">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="52a0b-271">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-271">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="52a0b-272">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-272">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="52a0b-273">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-273">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="52a0b-274">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-274">-BindingEvents</span></span><br><span data-ttu-id="52a0b-275">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-275">
        -CompressedFile</span></span><br><span data-ttu-id="52a0b-276">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-276">
        -DocumentEvents</span></span><br><span data-ttu-id="52a0b-277">
        - File</span><span class="sxs-lookup"><span data-stu-id="52a0b-277">
        - File</span></span><br><span data-ttu-id="52a0b-278">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-278">
        -ImageCoercion</span></span><br><span data-ttu-id="52a0b-279">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-279">
        -MatrixBindings</span></span><br><span data-ttu-id="52a0b-280">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-280">
        -MatrixCoercion</span></span><br><span data-ttu-id="52a0b-281">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-281">
        -PdfFile</span></span><br><span data-ttu-id="52a0b-282">
        - Sélection</span><span class="sxs-lookup"><span data-stu-id="52a0b-282">
        - Selection</span></span><br><span data-ttu-id="52a0b-283">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-283">
        - Settings</span></span><br><span data-ttu-id="52a0b-284">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-284">
        -TableBindings</span></span><br><span data-ttu-id="52a0b-285">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-285">
        -TableCoercion</span></span><br><span data-ttu-id="52a0b-286">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-286">
        -TextBindings</span></span><br><span data-ttu-id="52a0b-287">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-287">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="52a0b-288">Outlook</span><span class="sxs-lookup"><span data-stu-id="52a0b-288">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="52a0b-289">Plateforme</span><span class="sxs-lookup"><span data-stu-id="52a0b-289">Platform</span></span></th>
    <th><span data-ttu-id="52a0b-290">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="52a0b-290">Extension points</span></span></th>
    <th><span data-ttu-id="52a0b-291">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="52a0b-291">API requirement sets</span></span></th>
    <th><span data-ttu-id="52a0b-292"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="52a0b-292"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-293">Office Online</span><span class="sxs-lookup"><span data-stu-id="52a0b-293">Office Online</span></span></td>
    <td> <span data-ttu-id="52a0b-294">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="52a0b-294">- Mail Read</span></span><br><span data-ttu-id="52a0b-295">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="52a0b-295">
      - Mail Compose</span></span><br><span data-ttu-id="52a0b-296">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-296">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52a0b-297">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-297">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="52a0b-298">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-298">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="52a0b-299">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-299">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="52a0b-300">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-300">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="52a0b-301">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-301">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="52a0b-302">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-302">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="52a0b-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="52a0b-304">Non disponible</span><span class="sxs-lookup"><span data-stu-id="52a0b-304">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-305">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="52a0b-305">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="52a0b-306">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="52a0b-306">- Mail Read</span></span><br><span data-ttu-id="52a0b-307">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="52a0b-307">
      - Mail Compose</span></span><br><span data-ttu-id="52a0b-308">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-308">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52a0b-309">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-309">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="52a0b-310">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-310">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="52a0b-311">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-311">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="52a0b-312">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-312">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="52a0b-313">Non disponible</span><span class="sxs-lookup"><span data-stu-id="52a0b-313">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-314">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="52a0b-314">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="52a0b-315">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="52a0b-315">- Mail Read</span></span><br><span data-ttu-id="52a0b-316">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="52a0b-316">
      - Mail Compose</span></span><br><span data-ttu-id="52a0b-317">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-317">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="52a0b-318">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="52a0b-318">
      - Modules</span></span></td>
    <td> <span data-ttu-id="52a0b-319">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-319">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="52a0b-320">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-320">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="52a0b-321">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-321">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="52a0b-322">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-322">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="52a0b-323">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-323">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="52a0b-324">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-324">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="52a0b-325">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-325">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="52a0b-326">Non disponible</span><span class="sxs-lookup"><span data-stu-id="52a0b-326">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-327">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="52a0b-327">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="52a0b-328">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="52a0b-328">- Mail Read</span></span><br><span data-ttu-id="52a0b-329">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="52a0b-329">
      - Mail Compose</span></span><br><span data-ttu-id="52a0b-330">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-330">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="52a0b-331">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="52a0b-331">
      - Modules</span></span></td>
    <td> <span data-ttu-id="52a0b-332">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-332">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="52a0b-333">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-333">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="52a0b-334">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-334">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="52a0b-335">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-335">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="52a0b-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="52a0b-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="52a0b-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="52a0b-339">Non disponible</span><span class="sxs-lookup"><span data-stu-id="52a0b-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-340">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="52a0b-340">Office for iOS</span></span></td>
    <td> <span data-ttu-id="52a0b-341">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="52a0b-341">- Mail Read</span></span><br><span data-ttu-id="52a0b-342">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-342">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52a0b-343">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-343">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="52a0b-344">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-344">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="52a0b-345">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-345">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="52a0b-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="52a0b-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="52a0b-348">Non disponible</span><span class="sxs-lookup"><span data-stu-id="52a0b-348">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-349">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="52a0b-349">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="52a0b-350">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="52a0b-350">- Mail Read</span></span><br><span data-ttu-id="52a0b-351">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="52a0b-351">
      - Mail Compose</span></span><br><span data-ttu-id="52a0b-352">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-352">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52a0b-353">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-353">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="52a0b-354">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-354">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="52a0b-355">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-355">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="52a0b-356">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-356">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="52a0b-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="52a0b-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="52a0b-359">Non disponible</span><span class="sxs-lookup"><span data-stu-id="52a0b-359">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-360">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="52a0b-360">Office for Mac</span></span></td>
    <td> <span data-ttu-id="52a0b-361">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="52a0b-361">- Mail Read</span></span><br><span data-ttu-id="52a0b-362">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="52a0b-362">
      - Mail Compose</span></span><br><span data-ttu-id="52a0b-363">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-363">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52a0b-364">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-364">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="52a0b-365">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-365">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="52a0b-366">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-366">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="52a0b-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="52a0b-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="52a0b-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="52a0b-370">Non disponible</span><span class="sxs-lookup"><span data-stu-id="52a0b-370">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-371">Office pour Android</span><span class="sxs-lookup"><span data-stu-id="52a0b-371">Office for Android</span></span></td>
    <td> <span data-ttu-id="52a0b-372">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="52a0b-372">- Mail Read</span></span><br><span data-ttu-id="52a0b-373">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-373">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52a0b-374">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-374">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="52a0b-375">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-375">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="52a0b-376">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-376">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="52a0b-377">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-377">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="52a0b-378">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-378">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="52a0b-379">Non disponible</span><span class="sxs-lookup"><span data-stu-id="52a0b-379">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="52a0b-380">Word</span><span class="sxs-lookup"><span data-stu-id="52a0b-380">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="52a0b-381">Plateforme</span><span class="sxs-lookup"><span data-stu-id="52a0b-381">Platform</span></span></th>
    <th><span data-ttu-id="52a0b-382">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="52a0b-382">Extension points</span></span></th>
    <th><span data-ttu-id="52a0b-383">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="52a0b-383">API requirement sets</span></span></th>
    <th><span data-ttu-id="52a0b-384"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="52a0b-384"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-385">Office Online</span><span class="sxs-lookup"><span data-stu-id="52a0b-385">Office Online</span></span></td>
    <td> <span data-ttu-id="52a0b-386">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-386">- Taskpane</span></span><br><span data-ttu-id="52a0b-387">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-387">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52a0b-388">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-388">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="52a0b-389">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-389">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="52a0b-390">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-390">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="52a0b-391">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-391">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52a0b-392">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-392">-BindingEvents</span></span><br><span data-ttu-id="52a0b-393">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="52a0b-393">
         -</span></span><br><span data-ttu-id="52a0b-394">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-394">
         -DocumentEvents</span></span><br><span data-ttu-id="52a0b-395">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="52a0b-395">
         - File</span></span><br><span data-ttu-id="52a0b-396">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-396">
         -HtmlCoercion</span></span><br><span data-ttu-id="52a0b-397">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-397">
         -ImageCoercion</span></span><br><span data-ttu-id="52a0b-398">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-398">
         -MatrixBindings</span></span><br><span data-ttu-id="52a0b-399">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-399">
         -MatrixCoercion</span></span><br><span data-ttu-id="52a0b-400">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-400">
         -OoxmlCoercion</span></span><br><span data-ttu-id="52a0b-401">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-401">
         -PdfFile</span></span><br><span data-ttu-id="52a0b-402">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="52a0b-402">
         - Selection</span></span><br><span data-ttu-id="52a0b-403">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-403">
         - Settings</span></span><br><span data-ttu-id="52a0b-404">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-404">
         -TableBindings</span></span><br><span data-ttu-id="52a0b-405">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-405">
         -TableCoercion</span></span><br><span data-ttu-id="52a0b-406">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-406">
         -TextBindings</span></span><br><span data-ttu-id="52a0b-407">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-407">
         -TextCoercion</span></span><br><span data-ttu-id="52a0b-408">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-408">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-409">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="52a0b-409">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="52a0b-410">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-410">- Taskpane</span></span></td>
    <td> <span data-ttu-id="52a0b-411">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-411">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52a0b-412">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-412">-BindingEvents</span></span><br><span data-ttu-id="52a0b-413">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-413">
         -CompressedFile</span></span><br><span data-ttu-id="52a0b-414">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="52a0b-414">
         -</span></span><br><span data-ttu-id="52a0b-415">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-415">
         -DocumentEvents</span></span><br><span data-ttu-id="52a0b-416">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="52a0b-416">
         - File</span></span><br><span data-ttu-id="52a0b-417">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-417">
         -HtmlCoercion</span></span><br><span data-ttu-id="52a0b-418">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-418">
         -ImageCoercion</span></span><br><span data-ttu-id="52a0b-419">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-419">
         -MatrixBindings</span></span><br><span data-ttu-id="52a0b-420">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-420">
         -MatrixCoercion</span></span><br><span data-ttu-id="52a0b-421">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-421">
         -OoxmlCoercion</span></span><br><span data-ttu-id="52a0b-422">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-422">
         -PdfFile</span></span><br><span data-ttu-id="52a0b-423">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="52a0b-423">
         - Selection</span></span><br><span data-ttu-id="52a0b-424">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-424">
         - Settings</span></span><br><span data-ttu-id="52a0b-425">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-425">
         -TableBindings</span></span><br><span data-ttu-id="52a0b-426">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-426">
         -TableCoercion</span></span><br><span data-ttu-id="52a0b-427">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-427">
         -TextBindings</span></span><br><span data-ttu-id="52a0b-428">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-428">
         -TextCoercion</span></span><br><span data-ttu-id="52a0b-429">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-429">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-430">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="52a0b-430">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="52a0b-431">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-431">- Taskpane</span></span><br><span data-ttu-id="52a0b-432">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-432">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52a0b-433">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-433">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="52a0b-434">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-434">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="52a0b-435">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-435">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="52a0b-436">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-436">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52a0b-437">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-437">-BindingEvents</span></span><br><span data-ttu-id="52a0b-438">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-438">
         -CompressedFile</span></span><br><span data-ttu-id="52a0b-439">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="52a0b-439">
         -</span></span><br><span data-ttu-id="52a0b-440">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-440">
         -DocumentEvents</span></span><br><span data-ttu-id="52a0b-441">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="52a0b-441">
         - File</span></span><br><span data-ttu-id="52a0b-442">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-442">
         -HtmlCoercion</span></span><br><span data-ttu-id="52a0b-443">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-443">
         -ImageCoercion</span></span><br><span data-ttu-id="52a0b-444">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-444">
         -MatrixBindings</span></span><br><span data-ttu-id="52a0b-445">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-445">
         -MatrixCoercion</span></span><br><span data-ttu-id="52a0b-446">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-446">
         -OoxmlCoercion</span></span><br><span data-ttu-id="52a0b-447">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-447">
         -PdfFile</span></span><br><span data-ttu-id="52a0b-448">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="52a0b-448">
         - Selection</span></span><br><span data-ttu-id="52a0b-449">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-449">
         - Settings</span></span><br><span data-ttu-id="52a0b-450">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-450">
         -TableBindings</span></span><br><span data-ttu-id="52a0b-451">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-451">
         -TableCoercion</span></span><br><span data-ttu-id="52a0b-452">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-452">
         -TextBindings</span></span><br><span data-ttu-id="52a0b-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-453">
         -TextCoercion</span></span><br><span data-ttu-id="52a0b-454">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-454">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-455">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="52a0b-455">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="52a0b-456">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-456">- Taskpane</span></span><br><span data-ttu-id="52a0b-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52a0b-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="52a0b-459">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-459">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="52a0b-460">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-460">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="52a0b-461">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-461">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52a0b-462">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-462">-BindingEvents</span></span><br><span data-ttu-id="52a0b-463">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-463">
         -CompressedFile</span></span><br><span data-ttu-id="52a0b-464">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="52a0b-464">
         -</span></span><br><span data-ttu-id="52a0b-465">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-465">
         -DocumentEvents</span></span><br><span data-ttu-id="52a0b-466">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="52a0b-466">
         - File</span></span><br><span data-ttu-id="52a0b-467">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-467">
         -HtmlCoercion</span></span><br><span data-ttu-id="52a0b-468">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-468">
         -ImageCoercion</span></span><br><span data-ttu-id="52a0b-469">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-469">
         -MatrixBindings</span></span><br><span data-ttu-id="52a0b-470">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-470">
         -MatrixCoercion</span></span><br><span data-ttu-id="52a0b-471">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-471">
         -OoxmlCoercion</span></span><br><span data-ttu-id="52a0b-472">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-472">
         -PdfFile</span></span><br><span data-ttu-id="52a0b-473">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="52a0b-473">
         - Selection</span></span><br><span data-ttu-id="52a0b-474">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-474">
         - Settings</span></span><br><span data-ttu-id="52a0b-475">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-475">
         -TableBindings</span></span><br><span data-ttu-id="52a0b-476">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-476">
         -TableCoercion</span></span><br><span data-ttu-id="52a0b-477">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-477">
         -TextBindings</span></span><br><span data-ttu-id="52a0b-478">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-478">
         -TextCoercion</span></span><br><span data-ttu-id="52a0b-479">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-479">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-480">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="52a0b-480">Office for iOS</span></span></td>
    <td> <span data-ttu-id="52a0b-481">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-481">- Taskpane</span></span></td>
    <td> <span data-ttu-id="52a0b-482">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-482">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="52a0b-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="52a0b-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="52a0b-485">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="52a0b-485">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="52a0b-486">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-486">-BindingEvents</span></span><br><span data-ttu-id="52a0b-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-487">
         -CompressedFile</span></span><br><span data-ttu-id="52a0b-488">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="52a0b-488">
         -</span></span><br><span data-ttu-id="52a0b-489">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-489">
         -DocumentEvents</span></span><br><span data-ttu-id="52a0b-490">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="52a0b-490">
         - File</span></span><br><span data-ttu-id="52a0b-491">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-491">
         -HtmlCoercion</span></span><br><span data-ttu-id="52a0b-492">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-492">
         -ImageCoercion</span></span><br><span data-ttu-id="52a0b-493">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-493">
         -MatrixBindings</span></span><br><span data-ttu-id="52a0b-494">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-494">
         -MatrixCoercion</span></span><br><span data-ttu-id="52a0b-495">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-495">
         -OoxmlCoercion</span></span><br><span data-ttu-id="52a0b-496">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-496">
         -PdfFile</span></span><br><span data-ttu-id="52a0b-497">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="52a0b-497">
         - Selection</span></span><br><span data-ttu-id="52a0b-498">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-498">
         - Settings</span></span><br><span data-ttu-id="52a0b-499">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-499">
         -TableBindings</span></span><br><span data-ttu-id="52a0b-500">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-500">
         -TableCoercion</span></span><br><span data-ttu-id="52a0b-501">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-501">
         -TextBindings</span></span><br><span data-ttu-id="52a0b-502">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-502">
         -TextCoercion</span></span><br><span data-ttu-id="52a0b-503">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-503">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-504">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="52a0b-504">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="52a0b-505">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-505">- Taskpane</span></span><br><span data-ttu-id="52a0b-506">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-506">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52a0b-507">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-507">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="52a0b-508">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-508">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="52a0b-509">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-509">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="52a0b-510">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="52a0b-510">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="52a0b-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-511">-BindingEvents</span></span><br><span data-ttu-id="52a0b-512">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-512">
         -CompressedFile</span></span><br><span data-ttu-id="52a0b-513">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="52a0b-513">
         -</span></span><br><span data-ttu-id="52a0b-514">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-514">
         -DocumentEvents</span></span><br><span data-ttu-id="52a0b-515">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="52a0b-515">
         - File</span></span><br><span data-ttu-id="52a0b-516">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-516">
         -HtmlCoercion</span></span><br><span data-ttu-id="52a0b-517">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-517">
         -ImageCoercion</span></span><br><span data-ttu-id="52a0b-518">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-518">
         -MatrixBindings</span></span><br><span data-ttu-id="52a0b-519">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-519">
         -MatrixCoercion</span></span><br><span data-ttu-id="52a0b-520">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-520">
         -OoxmlCoercion</span></span><br><span data-ttu-id="52a0b-521">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-521">
         -PdfFile</span></span><br><span data-ttu-id="52a0b-522">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="52a0b-522">
         - Selection</span></span><br><span data-ttu-id="52a0b-523">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-523">
         - Settings</span></span><br><span data-ttu-id="52a0b-524">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-524">
         -TableBindings</span></span><br><span data-ttu-id="52a0b-525">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-525">
         -TableCoercion</span></span><br><span data-ttu-id="52a0b-526">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-526">
         -TextBindings</span></span><br><span data-ttu-id="52a0b-527">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-527">
         -TextCoercion</span></span><br><span data-ttu-id="52a0b-528">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-528">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-529">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="52a0b-529">Office for Mac</span></span></td>
    <td> <span data-ttu-id="52a0b-530">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-530">- Taskpane</span></span><br><span data-ttu-id="52a0b-531">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-531">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52a0b-532">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-532">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="52a0b-533">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-533">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="52a0b-534">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-534">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="52a0b-535">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="52a0b-535">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="52a0b-536">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-536">-BindingEvents</span></span><br><span data-ttu-id="52a0b-537">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-537">
         -CompressedFile</span></span><br><span data-ttu-id="52a0b-538">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="52a0b-538">
         -</span></span><br><span data-ttu-id="52a0b-539">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-539">
         -DocumentEvents</span></span><br><span data-ttu-id="52a0b-540">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="52a0b-540">
         - File</span></span><br><span data-ttu-id="52a0b-541">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-541">
         -HtmlCoercion</span></span><br><span data-ttu-id="52a0b-542">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-542">
         -ImageCoercion</span></span><br><span data-ttu-id="52a0b-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-543">
         -MatrixBindings</span></span><br><span data-ttu-id="52a0b-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-544">
         -MatrixCoercion</span></span><br><span data-ttu-id="52a0b-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-545">
         -OoxmlCoercion</span></span><br><span data-ttu-id="52a0b-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-546">
         -PdfFile</span></span><br><span data-ttu-id="52a0b-547">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="52a0b-547">
         - Selection</span></span><br><span data-ttu-id="52a0b-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-548">
         - Settings</span></span><br><span data-ttu-id="52a0b-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-549">
         -TableBindings</span></span><br><span data-ttu-id="52a0b-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-550">
         -TableCoercion</span></span><br><span data-ttu-id="52a0b-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52a0b-551">
         -TextBindings</span></span><br><span data-ttu-id="52a0b-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-552">
         -TextCoercion</span></span><br><span data-ttu-id="52a0b-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-553">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="52a0b-554">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="52a0b-554">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="52a0b-555">Plateforme</span><span class="sxs-lookup"><span data-stu-id="52a0b-555">Platform</span></span></th>
    <th><span data-ttu-id="52a0b-556">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="52a0b-556">Extension points</span></span></th>
    <th><span data-ttu-id="52a0b-557">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="52a0b-557">API requirement sets</span></span></th>
    <th><span data-ttu-id="52a0b-558"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="52a0b-558"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-559">Office Online</span><span class="sxs-lookup"><span data-stu-id="52a0b-559">Office Online</span></span></td>
    <td> <span data-ttu-id="52a0b-560">- Contenu</span><span class="sxs-lookup"><span data-stu-id="52a0b-560">- Content</span></span><br><span data-ttu-id="52a0b-561">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-561">
         - Taskpane</span></span><br><span data-ttu-id="52a0b-562">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-562">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52a0b-563">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-563">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52a0b-564">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="52a0b-564">-ActiveView</span></span><br><span data-ttu-id="52a0b-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-565">
         -CompressedFile</span></span><br><span data-ttu-id="52a0b-566">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-566">
         -DocumentEvents</span></span><br><span data-ttu-id="52a0b-567">
         - File</span><span class="sxs-lookup"><span data-stu-id="52a0b-567">
         - File</span></span><br><span data-ttu-id="52a0b-568">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-568">
         -ImageCoercion</span></span><br><span data-ttu-id="52a0b-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-569">
         -PdfFile</span></span><br><span data-ttu-id="52a0b-570">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="52a0b-570">
         - Selection</span></span><br><span data-ttu-id="52a0b-571">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-571">
         - Settings</span></span><br><span data-ttu-id="52a0b-572">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-572">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-573">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="52a0b-573">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="52a0b-574">- Contenu</span><span class="sxs-lookup"><span data-stu-id="52a0b-574">- Content</span></span><br><span data-ttu-id="52a0b-575">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-575">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="52a0b-576">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="52a0b-576">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="52a0b-577">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="52a0b-577">-ActiveView</span></span><br><span data-ttu-id="52a0b-578">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-578">
         -CompressedFile</span></span><br><span data-ttu-id="52a0b-579">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-579">
         -DocumentEvents</span></span><br><span data-ttu-id="52a0b-580">
         - File</span><span class="sxs-lookup"><span data-stu-id="52a0b-580">
         - File</span></span><br><span data-ttu-id="52a0b-581">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-581">
         -ImageCoercion</span></span><br><span data-ttu-id="52a0b-582">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-582">
         -PdfFile</span></span><br><span data-ttu-id="52a0b-583">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="52a0b-583">
         - Selection</span></span><br><span data-ttu-id="52a0b-584">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-584">
         - Settings</span></span><br><span data-ttu-id="52a0b-585">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-585">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-586">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="52a0b-586">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="52a0b-587">- Contenu</span><span class="sxs-lookup"><span data-stu-id="52a0b-587">- Content</span></span><br><span data-ttu-id="52a0b-588">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-588">
         - Taskpane</span></span><br><span data-ttu-id="52a0b-589">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-589">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52a0b-590">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-590">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52a0b-591">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="52a0b-591">-ActiveView</span></span><br><span data-ttu-id="52a0b-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-592">
         -CompressedFile</span></span><br><span data-ttu-id="52a0b-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-593">
         -DocumentEvents</span></span><br><span data-ttu-id="52a0b-594">
         - File</span><span class="sxs-lookup"><span data-stu-id="52a0b-594">
         - File</span></span><br><span data-ttu-id="52a0b-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-595">
         -ImageCoercion</span></span><br><span data-ttu-id="52a0b-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-596">
         -PdfFile</span></span><br><span data-ttu-id="52a0b-597">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="52a0b-597">
         - Selection</span></span><br><span data-ttu-id="52a0b-598">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-598">
         - Settings</span></span><br><span data-ttu-id="52a0b-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-599">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-600">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="52a0b-600">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="52a0b-601">- Contenu</span><span class="sxs-lookup"><span data-stu-id="52a0b-601">- Content</span></span><br><span data-ttu-id="52a0b-602">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-602">
         - Taskpane</span></span><br><span data-ttu-id="52a0b-603">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-603">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52a0b-604">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-604">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52a0b-605">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="52a0b-605">-ActiveView</span></span><br><span data-ttu-id="52a0b-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-606">
         -CompressedFile</span></span><br><span data-ttu-id="52a0b-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-607">
         -DocumentEvents</span></span><br><span data-ttu-id="52a0b-608">
         - File</span><span class="sxs-lookup"><span data-stu-id="52a0b-608">
         - File</span></span><br><span data-ttu-id="52a0b-609">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-609">
         -ImageCoercion</span></span><br><span data-ttu-id="52a0b-610">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-610">
         -PdfFile</span></span><br><span data-ttu-id="52a0b-611">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="52a0b-611">
         - Selection</span></span><br><span data-ttu-id="52a0b-612">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-612">
         - Settings</span></span><br><span data-ttu-id="52a0b-613">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-613">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-614">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="52a0b-614">Office for iOS</span></span></td>
    <td> <span data-ttu-id="52a0b-615">- Contenu</span><span class="sxs-lookup"><span data-stu-id="52a0b-615">- Content</span></span><br><span data-ttu-id="52a0b-616">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-616">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="52a0b-617">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-617">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="52a0b-618">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="52a0b-618">-ActiveView</span></span><br><span data-ttu-id="52a0b-619">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-619">
         -CompressedFile</span></span><br><span data-ttu-id="52a0b-620">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-620">
         -DocumentEvents</span></span><br><span data-ttu-id="52a0b-621">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="52a0b-621">
         - File</span></span><br><span data-ttu-id="52a0b-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-622">
         -PdfFile</span></span><br><span data-ttu-id="52a0b-623">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="52a0b-623">
         - Selection</span></span><br><span data-ttu-id="52a0b-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-624">
         - Settings</span></span><br><span data-ttu-id="52a0b-625">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-625">
         -TextCoercion</span></span><br><span data-ttu-id="52a0b-626">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-626">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-627">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="52a0b-627">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="52a0b-628">- Contenu</span><span class="sxs-lookup"><span data-stu-id="52a0b-628">- Content</span></span><br><span data-ttu-id="52a0b-629">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-629">
         - Taskpane</span></span><br><span data-ttu-id="52a0b-630">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-630">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52a0b-631">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-631">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52a0b-632">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="52a0b-632">-ActiveView</span></span><br><span data-ttu-id="52a0b-633">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-633">
         -CompressedFile</span></span><br><span data-ttu-id="52a0b-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-634">
         -DocumentEvents</span></span><br><span data-ttu-id="52a0b-635">
         - File</span><span class="sxs-lookup"><span data-stu-id="52a0b-635">
         - File</span></span><br><span data-ttu-id="52a0b-636">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-636">
         -ImageCoercion</span></span><br><span data-ttu-id="52a0b-637">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-637">
         -PdfFile</span></span><br><span data-ttu-id="52a0b-638">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="52a0b-638">
         - Selection</span></span><br><span data-ttu-id="52a0b-639">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-639">
         - Settings</span></span><br><span data-ttu-id="52a0b-640">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-640">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-641">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="52a0b-641">Office for Mac</span></span></td>
    <td> <span data-ttu-id="52a0b-642">- Contenu</span><span class="sxs-lookup"><span data-stu-id="52a0b-642">- Content</span></span><br><span data-ttu-id="52a0b-643">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-643">
         - Taskpane</span></span><br><span data-ttu-id="52a0b-644">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-644">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52a0b-645">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-645">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52a0b-646">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="52a0b-646">-ActiveView</span></span><br><span data-ttu-id="52a0b-647">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-647">
         -CompressedFile</span></span><br><span data-ttu-id="52a0b-648">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-648">
         -DocumentEvents</span></span><br><span data-ttu-id="52a0b-649">
         - File</span><span class="sxs-lookup"><span data-stu-id="52a0b-649">
         - File</span></span><br><span data-ttu-id="52a0b-650">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-650">
         -ImageCoercion</span></span><br><span data-ttu-id="52a0b-651">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52a0b-651">
         -PdfFile</span></span><br><span data-ttu-id="52a0b-652">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="52a0b-652">
         - Selection</span></span><br><span data-ttu-id="52a0b-653">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-653">
         - Settings</span></span><br><span data-ttu-id="52a0b-654">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-654">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="52a0b-655">OneNote</span><span class="sxs-lookup"><span data-stu-id="52a0b-655">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="52a0b-656">Plateforme</span><span class="sxs-lookup"><span data-stu-id="52a0b-656">Platform</span></span></th>
    <th><span data-ttu-id="52a0b-657">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="52a0b-657">Extension points</span></span></th>
    <th><span data-ttu-id="52a0b-658">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="52a0b-658">API requirement sets</span></span></th>
    <th><span data-ttu-id="52a0b-659"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="52a0b-659"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="52a0b-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="52a0b-660">Office Online</span></span></td>
    <td> <span data-ttu-id="52a0b-661">- Contenu</span><span class="sxs-lookup"><span data-stu-id="52a0b-661">- Content</span></span><br><span data-ttu-id="52a0b-662">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-662">
         - Taskpane</span></span><br><span data-ttu-id="52a0b-663">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-663">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52a0b-664">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-664">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="52a0b-665">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52a0b-665">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52a0b-666">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52a0b-666">-DocumentEvents</span></span><br><span data-ttu-id="52a0b-667">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-667">
         -HtmlCoercion</span></span><br><span data-ttu-id="52a0b-668">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-668">
         -ImageCoercion</span></span><br><span data-ttu-id="52a0b-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52a0b-669">
         - Settings</span></span><br><span data-ttu-id="52a0b-670">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52a0b-670">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="52a0b-671">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="52a0b-671">See also</span></span>

- [<span data-ttu-id="52a0b-672">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-672">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="52a0b-673">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="52a0b-673">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="52a0b-674">Ensembles de conditions requises des commandes de complément</span><span class="sxs-lookup"><span data-stu-id="52a0b-674">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="52a0b-675">Référence de l’interface API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="52a0b-675">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)
