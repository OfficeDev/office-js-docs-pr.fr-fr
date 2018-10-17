---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint et OneNote.
ms.date: 10/03/2018
ms.openlocfilehash: 39a80f322c282e29e6e8c4363f0c82522b33b75d
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579925"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="1f859-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="1f859-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="1f859-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d’un hôte Office spécifique, d’un ensemble de conditions requises, d’un membre d’API ou d’une version de l’API. Les tableaux suivants contiennent la plateforme disponible, les points d’extension, les ensembles d’API requis et les ensembles d’API courantes requis qui sont actuellement pris en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="1f859-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="1f859-p102">Si une cellule de tableau contient un astérisque (\*), cela signifie que nous travaillons dessus. Pour les ensembles de conditions requises pour Projet ou Access, voir [Ensembles de conditions requises communs à Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="1f859-p102">If a table cell contains an asterisk ( \* ), that means we’re working on it. For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="1f859-p103">Le numéro de build pour Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que les ensembles de conditions requises des API communes, ExcelApi 1.1 et WordApi 1.1.</span><span class="sxs-lookup"><span data-stu-id="1f859-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="1f859-110">Excel</span><span class="sxs-lookup"><span data-stu-id="1f859-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="1f859-111">Plateforme</span><span class="sxs-lookup"><span data-stu-id="1f859-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="1f859-112">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="1f859-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="1f859-113">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="1f859-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="1f859-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="1f859-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="1f859-115">Office Online</span></span></td>
    <td> <span data-ttu-id="1f859-116">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-116">- Taskpane</span></span><br><span data-ttu-id="1f859-117">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="1f859-117">
        - Content</span></span><br><span data-ttu-id="1f859-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="1f859-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="1f859-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1f859-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1f859-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1f859-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1f859-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1f859-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1f859-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1f859-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1f859-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1f859-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1f859-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="1f859-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1f859-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1f859-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1f859-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-128">
        -BindingEvents</span></span><br><span data-ttu-id="1f859-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-129">
        -CompressedFile</span></span><br><span data-ttu-id="1f859-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-130">
        -DocumentEvents</span></span><br><span data-ttu-id="1f859-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="1f859-131">
        - File</span></span><br><span data-ttu-id="1f859-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-132">
        -MatrixBindings</span></span><br><span data-ttu-id="1f859-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-133">
        -MatrixCoercion</span></span><br><span data-ttu-id="1f859-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1f859-134">
        - Selection</span></span><br><span data-ttu-id="1f859-135">
        - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-135">
        - Settings</span></span><br><span data-ttu-id="1f859-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-136">
        -TableBindings</span></span><br><span data-ttu-id="1f859-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-137">
        -TableCoercion</span></span><br><span data-ttu-id="1f859-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-138">
        -TextBindings</span></span><br><span data-ttu-id="1f859-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-139">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-140">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1f859-140">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="1f859-141">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-141">
        - Taskpane</span></span><br><span data-ttu-id="1f859-142">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="1f859-142">
        - Content</span></span></td>
    <td>  <span data-ttu-id="1f859-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1f859-144">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-144">
        -BindingEvents</span></span><br><span data-ttu-id="1f859-145">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-145">
        -CompressedFile</span></span><br><span data-ttu-id="1f859-146">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-146">
        -DocumentEvents</span></span><br><span data-ttu-id="1f859-147">
        - File</span><span class="sxs-lookup"><span data-stu-id="1f859-147">
        - File</span></span><br><span data-ttu-id="1f859-148">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-148">
        -ImageCoercion</span></span><br><span data-ttu-id="1f859-149">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-149">
        -MatrixBindings</span></span><br><span data-ttu-id="1f859-150">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-150">
        -MatrixCoercion</span></span><br><span data-ttu-id="1f859-151">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1f859-151">
        - Selection</span></span><br><span data-ttu-id="1f859-152">
        - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-152">
        - Settings</span></span><br><span data-ttu-id="1f859-153">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-153">
        -TableBindings</span></span><br><span data-ttu-id="1f859-154">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-154">
        -TableCoercion</span></span><br><span data-ttu-id="1f859-155">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-155">
        -TextBindings</span></span><br><span data-ttu-id="1f859-156">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-156">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-157">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1f859-157">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="1f859-158">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-158">- Taskpane</span></span><br><span data-ttu-id="1f859-159">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="1f859-159">
        - Content</span></span><br><span data-ttu-id="1f859-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1f859-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1f859-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1f859-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1f859-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1f859-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1f859-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1f859-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1f859-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1f859-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1f859-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1f859-167">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="1f859-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1f859-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1f859-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1f859-170">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-170">-BindingEvents</span></span><br><span data-ttu-id="1f859-171">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-171">
        -CompressedFile</span></span><br><span data-ttu-id="1f859-172">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-172">
        -DocumentEvents</span></span><br><span data-ttu-id="1f859-173">
        - File</span><span class="sxs-lookup"><span data-stu-id="1f859-173">
        - File</span></span><br><span data-ttu-id="1f859-174">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-174">
        -ImageCoercion</span></span><br><span data-ttu-id="1f859-175">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-175">
        -MatrixBindings</span></span><br><span data-ttu-id="1f859-176">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-176">
        -MatrixCoercion</span></span><br><span data-ttu-id="1f859-177">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1f859-177">
        - Selection</span></span><br><span data-ttu-id="1f859-178">
        - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-178">
        - Settings</span></span><br><span data-ttu-id="1f859-179">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-179">
        -TableBindings</span></span><br><span data-ttu-id="1f859-180">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-180">
        -TableCoercion</span></span><br><span data-ttu-id="1f859-181">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-181">
        -TextBindings</span></span><br><span data-ttu-id="1f859-182">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-182">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-183">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1f859-183">Office for Windows Desktop.</span></span></td>
    <td><span data-ttu-id="1f859-184">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-184">- Taskpane</span></span><br><span data-ttu-id="1f859-185">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="1f859-185">
        - Content</span></span><br><span data-ttu-id="1f859-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1f859-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1f859-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1f859-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1f859-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1f859-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1f859-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1f859-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1f859-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1f859-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1f859-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1f859-193">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="1f859-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1f859-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1f859-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1f859-196">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-196">-BindingEvents</span></span><br><span data-ttu-id="1f859-197">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-197">
        -CompressedFile</span></span><br><span data-ttu-id="1f859-198">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-198">
        -DocumentEvents</span></span><br><span data-ttu-id="1f859-199">
        - File</span><span class="sxs-lookup"><span data-stu-id="1f859-199">
        - File</span></span><br><span data-ttu-id="1f859-200">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-200">
        -ImageCoercion</span></span><br><span data-ttu-id="1f859-201">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-201">
        -MatrixBindings</span></span><br><span data-ttu-id="1f859-202">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-202">
        -MatrixCoercion</span></span><br><span data-ttu-id="1f859-203">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1f859-203">
        - Selection</span></span><br><span data-ttu-id="1f859-204">
        - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-204">
        - Settings</span></span><br><span data-ttu-id="1f859-205">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-205">
        -TableBindings</span></span><br><span data-ttu-id="1f859-206">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-206">
        -TableCoercion</span></span><br><span data-ttu-id="1f859-207">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-207">
        -TextBindings</span></span><br><span data-ttu-id="1f859-208">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-208">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-209">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="1f859-209">Office for iOS</span></span></td>
    <td><span data-ttu-id="1f859-210">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-210">- Taskpane</span></span><br><span data-ttu-id="1f859-211">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="1f859-211">
        - Content</span></span></td>
    <td><span data-ttu-id="1f859-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1f859-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1f859-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1f859-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1f859-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1f859-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1f859-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1f859-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1f859-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1f859-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1f859-218">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="1f859-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1f859-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1f859-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1f859-221">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-221">-BindingEvents</span></span><br><span data-ttu-id="1f859-222">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-222">
        -CompressedFile</span></span><br><span data-ttu-id="1f859-223">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-223">
        -DocumentEvents</span></span><br><span data-ttu-id="1f859-224">
        - File</span><span class="sxs-lookup"><span data-stu-id="1f859-224">
        - File</span></span><br><span data-ttu-id="1f859-225">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-225">
        -ImageCoercion</span></span><br><span data-ttu-id="1f859-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-226">
        -MatrixBindings</span></span><br><span data-ttu-id="1f859-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-227">
        -MatrixCoercion</span></span><br><span data-ttu-id="1f859-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1f859-228">
        - Selection</span></span><br><span data-ttu-id="1f859-229">
        - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-229">
        - Settings</span></span><br><span data-ttu-id="1f859-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-230">
        -TableBindings</span></span><br><span data-ttu-id="1f859-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-231">
        -TableCoercion</span></span><br><span data-ttu-id="1f859-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-232">
        -TextBindings</span></span><br><span data-ttu-id="1f859-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-233">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-234">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="1f859-234">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="1f859-235">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-235">- Taskpane</span></span><br><span data-ttu-id="1f859-236">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="1f859-236">
        - Content</span></span><br><span data-ttu-id="1f859-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1f859-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1f859-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1f859-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1f859-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1f859-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1f859-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1f859-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1f859-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1f859-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1f859-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1f859-244">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="1f859-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1f859-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1f859-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1f859-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-247">-BindingEvents</span></span><br><span data-ttu-id="1f859-248">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-248">
        -CompressedFile</span></span><br><span data-ttu-id="1f859-249">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-249">
        -DocumentEvents</span></span><br><span data-ttu-id="1f859-250">
        - File</span><span class="sxs-lookup"><span data-stu-id="1f859-250">
        - File</span></span><br><span data-ttu-id="1f859-251">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-251">
        -ImageCoercion</span></span><br><span data-ttu-id="1f859-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-252">
        -MatrixBindings</span></span><br><span data-ttu-id="1f859-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-253">
        -MatrixCoercion</span></span><br><span data-ttu-id="1f859-254">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1f859-254">
        -PdfFile</span></span><br><span data-ttu-id="1f859-255">
        - Sélection</span><span class="sxs-lookup"><span data-stu-id="1f859-255">
        - Selection</span></span><br><span data-ttu-id="1f859-256">
        - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-256">
        - Settings</span></span><br><span data-ttu-id="1f859-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-257">
        -TableBindings</span></span><br><span data-ttu-id="1f859-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-258">
        -TableCoercion</span></span><br><span data-ttu-id="1f859-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-259">
        -TextBindings</span></span><br><span data-ttu-id="1f859-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-260">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-261">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="1f859-261">Office for Mac</span></span></td>
    <td><span data-ttu-id="1f859-262">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-262">- Taskpane</span></span><br><span data-ttu-id="1f859-263">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="1f859-263">
        - Content</span></span><br><span data-ttu-id="1f859-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1f859-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1f859-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1f859-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1f859-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1f859-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1f859-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1f859-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1f859-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1f859-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1f859-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1f859-271">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="1f859-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1f859-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1f859-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1f859-274">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-274">-BindingEvents</span></span><br><span data-ttu-id="1f859-275">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-275">
        -CompressedFile</span></span><br><span data-ttu-id="1f859-276">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-276">
        -DocumentEvents</span></span><br><span data-ttu-id="1f859-277">
        - File</span><span class="sxs-lookup"><span data-stu-id="1f859-277">
        - File</span></span><br><span data-ttu-id="1f859-278">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-278">
        -ImageCoercion</span></span><br><span data-ttu-id="1f859-279">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-279">
        -MatrixBindings</span></span><br><span data-ttu-id="1f859-280">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-280">
        -MatrixCoercion</span></span><br><span data-ttu-id="1f859-281">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1f859-281">
        -PdfFile</span></span><br><span data-ttu-id="1f859-282">
        - Sélection</span><span class="sxs-lookup"><span data-stu-id="1f859-282">
        - Selection</span></span><br><span data-ttu-id="1f859-283">
        - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-283">
        - Settings</span></span><br><span data-ttu-id="1f859-284">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-284">
        -TableBindings</span></span><br><span data-ttu-id="1f859-285">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-285">
        -TableCoercion</span></span><br><span data-ttu-id="1f859-286">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-286">
        -TextBindings</span></span><br><span data-ttu-id="1f859-287">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-287">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="1f859-288">Outlook</span><span class="sxs-lookup"><span data-stu-id="1f859-288">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1f859-289">Plateforme</span><span class="sxs-lookup"><span data-stu-id="1f859-289">Platform</span></span></th>
    <th><span data-ttu-id="1f859-290">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="1f859-290">Extension points</span></span></th>
    <th><span data-ttu-id="1f859-291">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="1f859-291">API requirement sets</span></span></th>
    <th><span data-ttu-id="1f859-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="1f859-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-293">Office Online</span><span class="sxs-lookup"><span data-stu-id="1f859-293">Office Online</span></span></td>
    <td> <span data-ttu-id="1f859-294">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="1f859-294">- Mail Read</span></span><br><span data-ttu-id="1f859-295">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="1f859-295">
      - Mail Compose</span></span><br><span data-ttu-id="1f859-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1f859-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1f859-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1f859-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1f859-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1f859-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1f859-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1f859-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1f859-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1f859-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1f859-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1f859-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1f859-304">Non disponible</span><span class="sxs-lookup"><span data-stu-id="1f859-304">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-305">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1f859-305">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="1f859-306">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="1f859-306">- Mail Read</span></span><br><span data-ttu-id="1f859-307">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="1f859-307">
      - Mail Compose</span></span><br><span data-ttu-id="1f859-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1f859-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1f859-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1f859-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1f859-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1f859-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="1f859-313">Non disponible</span><span class="sxs-lookup"><span data-stu-id="1f859-313">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-314">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1f859-314">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="1f859-315">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="1f859-315">- Mail Read</span></span><br><span data-ttu-id="1f859-316">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="1f859-316">
      - Mail Compose</span></span><br><span data-ttu-id="1f859-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="1f859-318">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="1f859-318">
      - Modules</span></span></td>
    <td> <span data-ttu-id="1f859-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1f859-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1f859-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1f859-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1f859-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1f859-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1f859-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1f859-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1f859-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1f859-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1f859-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1f859-326">Non disponible</span><span class="sxs-lookup"><span data-stu-id="1f859-326">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-327">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1f859-327">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="1f859-328">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="1f859-328">- Mail Read</span></span><br><span data-ttu-id="1f859-329">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="1f859-329">
      - Mail Compose</span></span><br><span data-ttu-id="1f859-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="1f859-331">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="1f859-331">
      - Modules</span></span></td>
    <td> <span data-ttu-id="1f859-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1f859-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1f859-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1f859-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1f859-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1f859-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1f859-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1f859-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1f859-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1f859-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1f859-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1f859-339">Non disponible</span><span class="sxs-lookup"><span data-stu-id="1f859-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-340">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="1f859-340">Office for iOS</span></span></td>
    <td> <span data-ttu-id="1f859-341">- Lecture du courrier</span><span class="sxs-lookup"><span data-stu-id="1f859-341">- Mail Read</span></span><br><span data-ttu-id="1f859-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1f859-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1f859-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1f859-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1f859-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1f859-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1f859-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1f859-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="1f859-348">Non disponible</span><span class="sxs-lookup"><span data-stu-id="1f859-348">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-349">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="1f859-349">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="1f859-350">- Lecture du courrier</span><span class="sxs-lookup"><span data-stu-id="1f859-350">- Mail Read</span></span><br><span data-ttu-id="1f859-351">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="1f859-351">
      - Mail Compose</span></span><br><span data-ttu-id="1f859-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1f859-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1f859-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1f859-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1f859-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1f859-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1f859-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1f859-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1f859-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1f859-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="1f859-359">Non disponible</span><span class="sxs-lookup"><span data-stu-id="1f859-359">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-360">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="1f859-360">Office for Mac</span></span></td>
    <td> <span data-ttu-id="1f859-361">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="1f859-361">- Mail Read</span></span><br><span data-ttu-id="1f859-362">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="1f859-362">
      - Mail Compose</span></span><br><span data-ttu-id="1f859-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1f859-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1f859-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1f859-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1f859-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1f859-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1f859-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1f859-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1f859-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1f859-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1f859-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1f859-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1f859-371">Non disponible</span><span class="sxs-lookup"><span data-stu-id="1f859-371">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-372">Office pour Android</span><span class="sxs-lookup"><span data-stu-id="1f859-372">Office for Android</span></span></td>
    <td> <span data-ttu-id="1f859-373">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="1f859-373">- Mail Read</span></span><br><span data-ttu-id="1f859-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1f859-375">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-375">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1f859-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1f859-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1f859-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1f859-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1f859-379">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1f859-379">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="1f859-380">Non disponible</span><span class="sxs-lookup"><span data-stu-id="1f859-380">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="1f859-381">Word</span><span class="sxs-lookup"><span data-stu-id="1f859-381">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1f859-382">Plateforme</span><span class="sxs-lookup"><span data-stu-id="1f859-382">Platform</span></span></th>
    <th><span data-ttu-id="1f859-383">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="1f859-383">Extension points</span></span></th>
    <th><span data-ttu-id="1f859-384">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="1f859-384">API requirement sets</span></span></th>
    <th><span data-ttu-id="1f859-385"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="1f859-385"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-386">Office Online</span><span class="sxs-lookup"><span data-stu-id="1f859-386">Office Online</span></span></td>
    <td> <span data-ttu-id="1f859-387">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-387">- Taskpane</span></span><br><span data-ttu-id="1f859-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1f859-389">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-389">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1f859-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1f859-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1f859-392">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-392">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1f859-393">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-393">-BindingEvents</span></span><br><span data-ttu-id="1f859-394">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1f859-394">
         -</span></span><br><span data-ttu-id="1f859-395">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-395">
         -DocumentEvents</span></span><br><span data-ttu-id="1f859-396">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="1f859-396">
         - File</span></span><br><span data-ttu-id="1f859-397">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-397">
         -HtmlCoercion</span></span><br><span data-ttu-id="1f859-398">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-398">
         -ImageCoercion</span></span><br><span data-ttu-id="1f859-399">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-399">
         -MatrixBindings</span></span><br><span data-ttu-id="1f859-400">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-400">
         -MatrixCoercion</span></span><br><span data-ttu-id="1f859-401">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-401">
         -OoxmlCoercion</span></span><br><span data-ttu-id="1f859-402">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1f859-402">
         -PdfFile</span></span><br><span data-ttu-id="1f859-403">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="1f859-403">
         - Selection</span></span><br><span data-ttu-id="1f859-404">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-404">
         - Settings</span></span><br><span data-ttu-id="1f859-405">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-405">
         -TableBindings</span></span><br><span data-ttu-id="1f859-406">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-406">
         -TableCoercion</span></span><br><span data-ttu-id="1f859-407">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-407">
         -TextBindings</span></span><br><span data-ttu-id="1f859-408">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-408">
         -TextCoercion</span></span><br><span data-ttu-id="1f859-409">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1f859-409">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-410">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1f859-410">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="1f859-411">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-411">- Taskpane</span></span></td>
    <td> <span data-ttu-id="1f859-412">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-412">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1f859-413">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-413">-BindingEvents</span></span><br><span data-ttu-id="1f859-414">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-414">
         -CompressedFile</span></span><br><span data-ttu-id="1f859-415">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1f859-415">
         -</span></span><br><span data-ttu-id="1f859-416">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-416">
         -DocumentEvents</span></span><br><span data-ttu-id="1f859-417">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="1f859-417">
         - File</span></span><br><span data-ttu-id="1f859-418">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-418">
         -HtmlCoercion</span></span><br><span data-ttu-id="1f859-419">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-419">
         -ImageCoercion</span></span><br><span data-ttu-id="1f859-420">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-420">
         -MatrixBindings</span></span><br><span data-ttu-id="1f859-421">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-421">
         -MatrixCoercion</span></span><br><span data-ttu-id="1f859-422">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-422">
         -OoxmlCoercion</span></span><br><span data-ttu-id="1f859-423">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1f859-423">
         -PdfFile</span></span><br><span data-ttu-id="1f859-424">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="1f859-424">
         - Selection</span></span><br><span data-ttu-id="1f859-425">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-425">
         - Settings</span></span><br><span data-ttu-id="1f859-426">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-426">
         -TableBindings</span></span><br><span data-ttu-id="1f859-427">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-427">
         -TableCoercion</span></span><br><span data-ttu-id="1f859-428">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-428">
         -TextBindings</span></span><br><span data-ttu-id="1f859-429">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-429">
         -TextCoercion</span></span><br><span data-ttu-id="1f859-430">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1f859-430">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-431">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1f859-431">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="1f859-432">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-432">- Taskpane</span></span><br><span data-ttu-id="1f859-433">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-433">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1f859-434">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-434">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1f859-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1f859-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1f859-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1f859-438">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-438">-BindingEvents</span></span><br><span data-ttu-id="1f859-439">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-439">
         -CompressedFile</span></span><br><span data-ttu-id="1f859-440">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1f859-440">
         -</span></span><br><span data-ttu-id="1f859-441">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-441">
         -DocumentEvents</span></span><br><span data-ttu-id="1f859-442">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="1f859-442">
         - File</span></span><br><span data-ttu-id="1f859-443">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-443">
         -HtmlCoercion</span></span><br><span data-ttu-id="1f859-444">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-444">
         -ImageCoercion</span></span><br><span data-ttu-id="1f859-445">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-445">
         -MatrixBindings</span></span><br><span data-ttu-id="1f859-446">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-446">
         -MatrixCoercion</span></span><br><span data-ttu-id="1f859-447">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-447">
         -OoxmlCoercion</span></span><br><span data-ttu-id="1f859-448">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1f859-448">
         -PdfFile</span></span><br><span data-ttu-id="1f859-449">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="1f859-449">
         - Selection</span></span><br><span data-ttu-id="1f859-450">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-450">
         - Settings</span></span><br><span data-ttu-id="1f859-451">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-451">
         -TableBindings</span></span><br><span data-ttu-id="1f859-452">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-452">
         -TableCoercion</span></span><br><span data-ttu-id="1f859-453">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-453">
         -TextBindings</span></span><br><span data-ttu-id="1f859-454">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-454">
         -TextCoercion</span></span><br><span data-ttu-id="1f859-455">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1f859-455">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-456">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1f859-456">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="1f859-457">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-457">- Taskpane</span></span><br><span data-ttu-id="1f859-458">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-458">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1f859-459">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-459">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1f859-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1f859-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1f859-462">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-462">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1f859-463">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-463">-BindingEvents</span></span><br><span data-ttu-id="1f859-464">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-464">
         -CompressedFile</span></span><br><span data-ttu-id="1f859-465">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1f859-465">
         -</span></span><br><span data-ttu-id="1f859-466">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-466">
         -DocumentEvents</span></span><br><span data-ttu-id="1f859-467">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="1f859-467">
         - File</span></span><br><span data-ttu-id="1f859-468">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-468">
         -HtmlCoercion</span></span><br><span data-ttu-id="1f859-469">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-469">
         -ImageCoercion</span></span><br><span data-ttu-id="1f859-470">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-470">
         -MatrixBindings</span></span><br><span data-ttu-id="1f859-471">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-471">
         -MatrixCoercion</span></span><br><span data-ttu-id="1f859-472">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-472">
         -OoxmlCoercion</span></span><br><span data-ttu-id="1f859-473">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1f859-473">
         -PdfFile</span></span><br><span data-ttu-id="1f859-474">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="1f859-474">
         - Selection</span></span><br><span data-ttu-id="1f859-475">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-475">
         - Settings</span></span><br><span data-ttu-id="1f859-476">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-476">
         -TableBindings</span></span><br><span data-ttu-id="1f859-477">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-477">
         -TableCoercion</span></span><br><span data-ttu-id="1f859-478">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-478">
         -TextBindings</span></span><br><span data-ttu-id="1f859-479">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-479">
         -TextCoercion</span></span><br><span data-ttu-id="1f859-480">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1f859-480">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-481">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="1f859-481">Office for iOS</span></span></td>
    <td> <span data-ttu-id="1f859-482">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-482">- Taskpane</span></span></td>
    <td> <span data-ttu-id="1f859-483">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-483">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1f859-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1f859-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1f859-486">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="1f859-486">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="1f859-487">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-487">-BindingEvents</span></span><br><span data-ttu-id="1f859-488">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-488">
         -CompressedFile</span></span><br><span data-ttu-id="1f859-489">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1f859-489">
         -</span></span><br><span data-ttu-id="1f859-490">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-490">
         -DocumentEvents</span></span><br><span data-ttu-id="1f859-491">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="1f859-491">
         - File</span></span><br><span data-ttu-id="1f859-492">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-492">
         -HtmlCoercion</span></span><br><span data-ttu-id="1f859-493">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-493">
         -ImageCoercion</span></span><br><span data-ttu-id="1f859-494">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-494">
         -MatrixBindings</span></span><br><span data-ttu-id="1f859-495">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-495">
         -MatrixCoercion</span></span><br><span data-ttu-id="1f859-496">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-496">
         -OoxmlCoercion</span></span><br><span data-ttu-id="1f859-497">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1f859-497">
         -PdfFile</span></span><br><span data-ttu-id="1f859-498">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="1f859-498">
         - Selection</span></span><br><span data-ttu-id="1f859-499">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-499">
         - Settings</span></span><br><span data-ttu-id="1f859-500">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-500">
         -TableBindings</span></span><br><span data-ttu-id="1f859-501">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-501">
         -TableCoercion</span></span><br><span data-ttu-id="1f859-502">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-502">
         -TextBindings</span></span><br><span data-ttu-id="1f859-503">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-503">
         -TextCoercion</span></span><br><span data-ttu-id="1f859-504">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1f859-504">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-505">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="1f859-505">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="1f859-506">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-506">- Taskpane</span></span><br><span data-ttu-id="1f859-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1f859-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1f859-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1f859-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1f859-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="1f859-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="1f859-512">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-512">-BindingEvents</span></span><br><span data-ttu-id="1f859-513">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-513">
         -CompressedFile</span></span><br><span data-ttu-id="1f859-514">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1f859-514">
         -</span></span><br><span data-ttu-id="1f859-515">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-515">
         -DocumentEvents</span></span><br><span data-ttu-id="1f859-516">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="1f859-516">
         - File</span></span><br><span data-ttu-id="1f859-517">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-517">
         -HtmlCoercion</span></span><br><span data-ttu-id="1f859-518">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-518">
         -ImageCoercion</span></span><br><span data-ttu-id="1f859-519">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-519">
         -MatrixBindings</span></span><br><span data-ttu-id="1f859-520">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-520">
         -MatrixCoercion</span></span><br><span data-ttu-id="1f859-521">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-521">
         -OoxmlCoercion</span></span><br><span data-ttu-id="1f859-522">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1f859-522">
         -PdfFile</span></span><br><span data-ttu-id="1f859-523">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="1f859-523">
         - Selection</span></span><br><span data-ttu-id="1f859-524">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-524">
         - Settings</span></span><br><span data-ttu-id="1f859-525">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-525">
         -TableBindings</span></span><br><span data-ttu-id="1f859-526">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-526">
         -TableCoercion</span></span><br><span data-ttu-id="1f859-527">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-527">
         -TextBindings</span></span><br><span data-ttu-id="1f859-528">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-528">
         -TextCoercion</span></span><br><span data-ttu-id="1f859-529">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1f859-529">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-530">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="1f859-530">Office for Mac</span></span></td>
    <td> <span data-ttu-id="1f859-531">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-531">- Taskpane</span></span><br><span data-ttu-id="1f859-532">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-532">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1f859-533">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-533">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1f859-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1f859-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1f859-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1f859-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1f859-536">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="1f859-536">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="1f859-537">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-537">-BindingEvents</span></span><br><span data-ttu-id="1f859-538">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-538">
         -CompressedFile</span></span><br><span data-ttu-id="1f859-539">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1f859-539">
         -</span></span><br><span data-ttu-id="1f859-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-540">
         -DocumentEvents</span></span><br><span data-ttu-id="1f859-541">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="1f859-541">
         - File</span></span><br><span data-ttu-id="1f859-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-542">
         -HtmlCoercion</span></span><br><span data-ttu-id="1f859-543">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-543">
         -ImageCoercion</span></span><br><span data-ttu-id="1f859-544">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-544">
         -MatrixBindings</span></span><br><span data-ttu-id="1f859-545">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-545">
         -MatrixCoercion</span></span><br><span data-ttu-id="1f859-546">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-546">
         -OoxmlCoercion</span></span><br><span data-ttu-id="1f859-547">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1f859-547">
         -PdfFile</span></span><br><span data-ttu-id="1f859-548">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="1f859-548">
         - Selection</span></span><br><span data-ttu-id="1f859-549">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-549">
         - Settings</span></span><br><span data-ttu-id="1f859-550">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-550">
         -TableBindings</span></span><br><span data-ttu-id="1f859-551">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-551">
         -TableCoercion</span></span><br><span data-ttu-id="1f859-552">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1f859-552">
         -TextBindings</span></span><br><span data-ttu-id="1f859-553">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-553">
         -TextCoercion</span></span><br><span data-ttu-id="1f859-554">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1f859-554">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="1f859-555">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="1f859-555">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1f859-556">Plateforme</span><span class="sxs-lookup"><span data-stu-id="1f859-556">Platform</span></span></th>
    <th><span data-ttu-id="1f859-557">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="1f859-557">Extension points</span></span></th>
    <th><span data-ttu-id="1f859-558">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="1f859-558">API requirement sets</span></span></th>
    <th><span data-ttu-id="1f859-559"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="1f859-559"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-560">Office Online</span><span class="sxs-lookup"><span data-stu-id="1f859-560">Office Online</span></span></td>
    <td> <span data-ttu-id="1f859-561">- Contenu</span><span class="sxs-lookup"><span data-stu-id="1f859-561">- Content</span></span><br><span data-ttu-id="1f859-562">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-562">
         - Taskpane</span></span><br><span data-ttu-id="1f859-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1f859-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1f859-565">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1f859-565">-ActiveView</span></span><br><span data-ttu-id="1f859-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-566">
         -CompressedFile</span></span><br><span data-ttu-id="1f859-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-567">
         -DocumentEvents</span></span><br><span data-ttu-id="1f859-568">
         - File</span><span class="sxs-lookup"><span data-stu-id="1f859-568">
         - File</span></span><br><span data-ttu-id="1f859-569">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-569">
         -ImageCoercion</span></span><br><span data-ttu-id="1f859-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1f859-570">
         -PdfFile</span></span><br><span data-ttu-id="1f859-571">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="1f859-571">
         - Selection</span></span><br><span data-ttu-id="1f859-572">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-572">
         - Settings</span></span><br><span data-ttu-id="1f859-573">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-573">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-574">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1f859-574">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="1f859-575">- Contenu</span><span class="sxs-lookup"><span data-stu-id="1f859-575">- Content</span></span><br><span data-ttu-id="1f859-576">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-576">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="1f859-577">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="1f859-577">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="1f859-578">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1f859-578">-ActiveView</span></span><br><span data-ttu-id="1f859-579">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-579">
         -CompressedFile</span></span><br><span data-ttu-id="1f859-580">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-580">
         -DocumentEvents</span></span><br><span data-ttu-id="1f859-581">
         - File</span><span class="sxs-lookup"><span data-stu-id="1f859-581">
         - File</span></span><br><span data-ttu-id="1f859-582">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-582">
         -ImageCoercion</span></span><br><span data-ttu-id="1f859-583">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1f859-583">
         -PdfFile</span></span><br><span data-ttu-id="1f859-584">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="1f859-584">
         - Selection</span></span><br><span data-ttu-id="1f859-585">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-585">
         - Settings</span></span><br><span data-ttu-id="1f859-586">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-586">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-587">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1f859-587">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="1f859-588">- Contenu</span><span class="sxs-lookup"><span data-stu-id="1f859-588">- Content</span></span><br><span data-ttu-id="1f859-589">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-589">
         - Taskpane</span></span><br><span data-ttu-id="1f859-590">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-590">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1f859-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1f859-592">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1f859-592">-ActiveView</span></span><br><span data-ttu-id="1f859-593">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-593">
         -CompressedFile</span></span><br><span data-ttu-id="1f859-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-594">
         -DocumentEvents</span></span><br><span data-ttu-id="1f859-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="1f859-595">
         - File</span></span><br><span data-ttu-id="1f859-596">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-596">
         -ImageCoercion</span></span><br><span data-ttu-id="1f859-597">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1f859-597">
         -PdfFile</span></span><br><span data-ttu-id="1f859-598">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="1f859-598">
         - Selection</span></span><br><span data-ttu-id="1f859-599">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-599">
         - Settings</span></span><br><span data-ttu-id="1f859-600">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-600">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-601">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1f859-601">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="1f859-602">- Contenu</span><span class="sxs-lookup"><span data-stu-id="1f859-602">- Content</span></span><br><span data-ttu-id="1f859-603">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-603">
         - Taskpane</span></span><br><span data-ttu-id="1f859-604">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-604">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1f859-605">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-605">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1f859-606">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1f859-606">-ActiveView</span></span><br><span data-ttu-id="1f859-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-607">
         -CompressedFile</span></span><br><span data-ttu-id="1f859-608">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-608">
         -DocumentEvents</span></span><br><span data-ttu-id="1f859-609">
         - File</span><span class="sxs-lookup"><span data-stu-id="1f859-609">
         - File</span></span><br><span data-ttu-id="1f859-610">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-610">
         -ImageCoercion</span></span><br><span data-ttu-id="1f859-611">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1f859-611">
         -PdfFile</span></span><br><span data-ttu-id="1f859-612">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="1f859-612">
         - Selection</span></span><br><span data-ttu-id="1f859-613">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-613">
         - Settings</span></span><br><span data-ttu-id="1f859-614">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-614">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-615">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="1f859-615">Office for iOS</span></span></td>
    <td> <span data-ttu-id="1f859-616">- Contenu</span><span class="sxs-lookup"><span data-stu-id="1f859-616">- Content</span></span><br><span data-ttu-id="1f859-617">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-617">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="1f859-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="1f859-619">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1f859-619">-ActiveView</span></span><br><span data-ttu-id="1f859-620">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-620">
         -CompressedFile</span></span><br><span data-ttu-id="1f859-621">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-621">
         -DocumentEvents</span></span><br><span data-ttu-id="1f859-622">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="1f859-622">
         - File</span></span><br><span data-ttu-id="1f859-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1f859-623">
         -PdfFile</span></span><br><span data-ttu-id="1f859-624">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="1f859-624">
         - Selection</span></span><br><span data-ttu-id="1f859-625">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-625">
         - Settings</span></span><br><span data-ttu-id="1f859-626">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-626">
         -TextCoercion</span></span><br><span data-ttu-id="1f859-627">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-627">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-628">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="1f859-628">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="1f859-629">- Contenu</span><span class="sxs-lookup"><span data-stu-id="1f859-629">- Content</span></span><br><span data-ttu-id="1f859-630">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-630">
         - Taskpane</span></span><br><span data-ttu-id="1f859-631">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-631">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1f859-632">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-632">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1f859-633">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1f859-633">-ActiveView</span></span><br><span data-ttu-id="1f859-634">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-634">
         -CompressedFile</span></span><br><span data-ttu-id="1f859-635">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-635">
         -DocumentEvents</span></span><br><span data-ttu-id="1f859-636">
         - File</span><span class="sxs-lookup"><span data-stu-id="1f859-636">
         - File</span></span><br><span data-ttu-id="1f859-637">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-637">
         -ImageCoercion</span></span><br><span data-ttu-id="1f859-638">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1f859-638">
         -PdfFile</span></span><br><span data-ttu-id="1f859-639">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="1f859-639">
         - Selection</span></span><br><span data-ttu-id="1f859-640">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-640">
         - Settings</span></span><br><span data-ttu-id="1f859-641">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-641">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-642">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="1f859-642">Office for Mac</span></span></td>
    <td> <span data-ttu-id="1f859-643">- Contenu</span><span class="sxs-lookup"><span data-stu-id="1f859-643">- Content</span></span><br><span data-ttu-id="1f859-644">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-644">
         - Taskpane</span></span><br><span data-ttu-id="1f859-645">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1f859-645">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1f859-646">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-646">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1f859-647">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1f859-647">-ActiveView</span></span><br><span data-ttu-id="1f859-648">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1f859-648">
         -CompressedFile</span></span><br><span data-ttu-id="1f859-649">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-649">
         -DocumentEvents</span></span><br><span data-ttu-id="1f859-650">
         - File</span><span class="sxs-lookup"><span data-stu-id="1f859-650">
         - File</span></span><br><span data-ttu-id="1f859-651">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-651">
         -ImageCoercion</span></span><br><span data-ttu-id="1f859-652">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1f859-652">
         -PdfFile</span></span><br><span data-ttu-id="1f859-653">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="1f859-653">
         - Selection</span></span><br><span data-ttu-id="1f859-654">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-654">
         - Settings</span></span><br><span data-ttu-id="1f859-655">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-655">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="1f859-656">OneNote</span><span class="sxs-lookup"><span data-stu-id="1f859-656">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1f859-657">Plateforme</span><span class="sxs-lookup"><span data-stu-id="1f859-657">Platform</span></span></th>
    <th><span data-ttu-id="1f859-658">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="1f859-658">Extension points</span></span></th>
    <th><span data-ttu-id="1f859-659">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="1f859-659">API requirement sets</span></span></th>
    <th><span data-ttu-id="1f859-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="1f859-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="1f859-661">Office Online</span><span class="sxs-lookup"><span data-stu-id="1f859-661">Office Online</span></span></td>
    <td> <span data-ttu-id="1f859-662">- Contenu</span><span class="sxs-lookup"><span data-stu-id="1f859-662">- Content</span></span><br><span data-ttu-id="1f859-663">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="1f859-663">
         - Taskpane</span></span><br><span data-ttu-id="1f859-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de compléments</a></span><span class="sxs-lookup"><span data-stu-id="1f859-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1f859-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="1f859-666">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1f859-666">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1f859-667">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1f859-667">-DocumentEvents</span></span><br><span data-ttu-id="1f859-668">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-668">
         -HtmlCoercion</span></span><br><span data-ttu-id="1f859-669">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-669">
         -ImageCoercion</span></span><br><span data-ttu-id="1f859-670">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="1f859-670">
         - Settings</span></span><br><span data-ttu-id="1f859-671">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1f859-671">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="1f859-672">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="1f859-672">See also</span></span>

- [<span data-ttu-id="1f859-673">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="1f859-673">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="1f859-674">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="1f859-674">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="1f859-675">Ensembles de conditions requises des commandes de complément</span><span class="sxs-lookup"><span data-stu-id="1f859-675">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="1f859-676">Référence de l’interface API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="1f859-676">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
