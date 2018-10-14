---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint et OneNote.
ms.date: 10/03/2018
ms.openlocfilehash: bc7ac5c97c041a546c160c05cffc2c80db1ff1b1
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506349"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="15618-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="15618-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="15618-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d’un hôte Office spécifique, d’un ensemble de conditions requises, d’un membre d’API ou d’une version de l’API. Les tableaux suivants contiennent la plateforme disponible, les points d’extension, les ensembles d’API requis et les ensembles d’API courantes requis qui sont actuellement pris en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="15618-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="15618-p102">Si une cellule de tableau contient un astérisque (\*), cela signifie que nous travaillons dessus. Pour les ensembles de conditions requises pour Projet ou Access, voir [Ensembles de conditions requises communs à Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="15618-p102">If a table cell contains an asterisk ( \* ), that means we’re working on it. For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="15618-p103">Le numéro de build pour Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que les ensembles de conditions requises des API communes, ExcelApi 1.1 et WordApi 1.1.</span><span class="sxs-lookup"><span data-stu-id="15618-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="15618-110">Excel</span><span class="sxs-lookup"><span data-stu-id="15618-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="15618-111">Plateforme</span><span class="sxs-lookup"><span data-stu-id="15618-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="15618-112">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="15618-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="15618-113">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="15618-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="15618-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="15618-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="15618-115">Office Online</span></span></td>
    <td> <span data-ttu-id="15618-116">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-116">- Taskpane</span></span><br><span data-ttu-id="15618-117">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="15618-117">
        - Content</span></span><br><span data-ttu-id="15618-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="15618-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="15618-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="15618-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="15618-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="15618-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="15618-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="15618-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="15618-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="15618-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="15618-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="15618-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="15618-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="15618-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="15618-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="15618-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="15618-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="15618-128">
        -BindingEvents</span></span><br><span data-ttu-id="15618-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-129">
        -CompressedFile</span></span><br><span data-ttu-id="15618-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-130">
        -DocumentEvents</span></span><br><span data-ttu-id="15618-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="15618-131">
        - File</span></span><br><span data-ttu-id="15618-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="15618-132">
        -MatrixBindings</span></span><br><span data-ttu-id="15618-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-133">
        -MatrixCoercion</span></span><br><span data-ttu-id="15618-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="15618-134">
        - Selection</span></span><br><span data-ttu-id="15618-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-135">
        - Settings</span></span><br><span data-ttu-id="15618-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="15618-136">
        -TableBindings</span></span><br><span data-ttu-id="15618-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-137">
        -TableCoercion</span></span><br><span data-ttu-id="15618-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="15618-138">
        -TextBindings</span></span><br><span data-ttu-id="15618-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-139">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-140">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="15618-140">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="15618-141">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-141">
        - Taskpane</span></span><br><span data-ttu-id="15618-142">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="15618-142">
        - Content</span></span></td>
    <td>  <span data-ttu-id="15618-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="15618-144">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="15618-144">
        -BindingEvents</span></span><br><span data-ttu-id="15618-145">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-145">
        -CompressedFile</span></span><br><span data-ttu-id="15618-146">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-146">
        -DocumentEvents</span></span><br><span data-ttu-id="15618-147">
        - File</span><span class="sxs-lookup"><span data-stu-id="15618-147">
        - File</span></span><br><span data-ttu-id="15618-148">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-148">
        -ImageCoercion</span></span><br><span data-ttu-id="15618-149">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="15618-149">
        -MatrixBindings</span></span><br><span data-ttu-id="15618-150">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-150">
        -MatrixCoercion</span></span><br><span data-ttu-id="15618-151">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="15618-151">
        - Selection</span></span><br><span data-ttu-id="15618-152">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-152">
        - Settings</span></span><br><span data-ttu-id="15618-153">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="15618-153">
        -TableBindings</span></span><br><span data-ttu-id="15618-154">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-154">
        -TableCoercion</span></span><br><span data-ttu-id="15618-155">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="15618-155">
        -TextBindings</span></span><br><span data-ttu-id="15618-156">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-156">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-157">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="15618-157">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="15618-158">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-158">- Taskpane</span></span><br><span data-ttu-id="15618-159">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="15618-159">
        - Content</span></span><br><span data-ttu-id="15618-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="15618-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="15618-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="15618-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="15618-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="15618-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="15618-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="15618-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="15618-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="15618-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="15618-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="15618-167">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="15618-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="15618-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="15618-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="15618-170">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="15618-170">-BindingEvents</span></span><br><span data-ttu-id="15618-171">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-171">
        -CompressedFile</span></span><br><span data-ttu-id="15618-172">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-172">
        -DocumentEvents</span></span><br><span data-ttu-id="15618-173">
        - File</span><span class="sxs-lookup"><span data-stu-id="15618-173">
        - File</span></span><br><span data-ttu-id="15618-174">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-174">
        -ImageCoercion</span></span><br><span data-ttu-id="15618-175">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="15618-175">
        -MatrixBindings</span></span><br><span data-ttu-id="15618-176">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-176">
        -MatrixCoercion</span></span><br><span data-ttu-id="15618-177">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="15618-177">
        - Selection</span></span><br><span data-ttu-id="15618-178">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-178">
        - Settings</span></span><br><span data-ttu-id="15618-179">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="15618-179">
        -TableBindings</span></span><br><span data-ttu-id="15618-180">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-180">
        -TableCoercion</span></span><br><span data-ttu-id="15618-181">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="15618-181">
        -TextBindings</span></span><br><span data-ttu-id="15618-182">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-182">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-183">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="15618-183">Office for Windows Desktop.</span></span></td>
    <td><span data-ttu-id="15618-184">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-184">- Taskpane</span></span><br><span data-ttu-id="15618-185">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="15618-185">
        - Content</span></span><br><span data-ttu-id="15618-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="15618-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="15618-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="15618-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="15618-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="15618-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="15618-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="15618-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="15618-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="15618-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="15618-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="15618-193">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="15618-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="15618-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="15618-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="15618-196">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="15618-196">-BindingEvents</span></span><br><span data-ttu-id="15618-197">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-197">
        -CompressedFile</span></span><br><span data-ttu-id="15618-198">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-198">
        -DocumentEvents</span></span><br><span data-ttu-id="15618-199">
        - File</span><span class="sxs-lookup"><span data-stu-id="15618-199">
        - File</span></span><br><span data-ttu-id="15618-200">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-200">
        -ImageCoercion</span></span><br><span data-ttu-id="15618-201">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="15618-201">
        -MatrixBindings</span></span><br><span data-ttu-id="15618-202">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-202">
        -MatrixCoercion</span></span><br><span data-ttu-id="15618-203">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="15618-203">
        - Selection</span></span><br><span data-ttu-id="15618-204">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-204">
        - Settings</span></span><br><span data-ttu-id="15618-205">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="15618-205">
        -TableBindings</span></span><br><span data-ttu-id="15618-206">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-206">
        -TableCoercion</span></span><br><span data-ttu-id="15618-207">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="15618-207">
        -TextBindings</span></span><br><span data-ttu-id="15618-208">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-208">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-209">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="15618-209">Office for iOS</span></span></td>
    <td><span data-ttu-id="15618-210">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-210">- Taskpane</span></span><br><span data-ttu-id="15618-211">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="15618-211">
        - Content</span></span></td>
    <td><span data-ttu-id="15618-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="15618-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="15618-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="15618-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="15618-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="15618-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="15618-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="15618-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="15618-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="15618-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="15618-218">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="15618-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="15618-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="15618-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="15618-221">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="15618-221">-BindingEvents</span></span><br><span data-ttu-id="15618-222">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-222">
        -CompressedFile</span></span><br><span data-ttu-id="15618-223">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-223">
        -DocumentEvents</span></span><br><span data-ttu-id="15618-224">
        - File</span><span class="sxs-lookup"><span data-stu-id="15618-224">
        - File</span></span><br><span data-ttu-id="15618-225">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-225">
        -ImageCoercion</span></span><br><span data-ttu-id="15618-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="15618-226">
        -MatrixBindings</span></span><br><span data-ttu-id="15618-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-227">
        -MatrixCoercion</span></span><br><span data-ttu-id="15618-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="15618-228">
        - Selection</span></span><br><span data-ttu-id="15618-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-229">
        - Settings</span></span><br><span data-ttu-id="15618-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="15618-230">
        -TableBindings</span></span><br><span data-ttu-id="15618-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-231">
        -TableCoercion</span></span><br><span data-ttu-id="15618-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="15618-232">
        -TextBindings</span></span><br><span data-ttu-id="15618-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-233">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-234">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="15618-234">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="15618-235">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-235">- Taskpane</span></span><br><span data-ttu-id="15618-236">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="15618-236">
        - Content</span></span><br><span data-ttu-id="15618-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="15618-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="15618-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="15618-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="15618-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="15618-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="15618-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="15618-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="15618-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="15618-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="15618-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="15618-244">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="15618-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="15618-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="15618-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="15618-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="15618-247">-BindingEvents</span></span><br><span data-ttu-id="15618-248">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-248">
        -CompressedFile</span></span><br><span data-ttu-id="15618-249">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-249">
        -DocumentEvents</span></span><br><span data-ttu-id="15618-250">
        - File</span><span class="sxs-lookup"><span data-stu-id="15618-250">
        - File</span></span><br><span data-ttu-id="15618-251">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-251">
        -ImageCoercion</span></span><br><span data-ttu-id="15618-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="15618-252">
        -MatrixBindings</span></span><br><span data-ttu-id="15618-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-253">
        -MatrixCoercion</span></span><br><span data-ttu-id="15618-254">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="15618-254">
        -PdfFile</span></span><br><span data-ttu-id="15618-255">
        - Sélection</span><span class="sxs-lookup"><span data-stu-id="15618-255">
        - Selection</span></span><br><span data-ttu-id="15618-256">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-256">
        - Settings</span></span><br><span data-ttu-id="15618-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="15618-257">
        -TableBindings</span></span><br><span data-ttu-id="15618-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-258">
        -TableCoercion</span></span><br><span data-ttu-id="15618-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="15618-259">
        -TextBindings</span></span><br><span data-ttu-id="15618-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-260">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-261">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="15618-261">Office for Mac</span></span></td>
    <td><span data-ttu-id="15618-262">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-262">- Taskpane</span></span><br><span data-ttu-id="15618-263">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="15618-263">
        - Content</span></span><br><span data-ttu-id="15618-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="15618-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="15618-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="15618-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="15618-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="15618-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="15618-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="15618-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="15618-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="15618-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="15618-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="15618-271">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="15618-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="15618-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="15618-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="15618-274">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="15618-274">-BindingEvents</span></span><br><span data-ttu-id="15618-275">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-275">
        -CompressedFile</span></span><br><span data-ttu-id="15618-276">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-276">
        -DocumentEvents</span></span><br><span data-ttu-id="15618-277">
        - File</span><span class="sxs-lookup"><span data-stu-id="15618-277">
        - File</span></span><br><span data-ttu-id="15618-278">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-278">
        -ImageCoercion</span></span><br><span data-ttu-id="15618-279">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="15618-279">
        -MatrixBindings</span></span><br><span data-ttu-id="15618-280">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-280">
        -MatrixCoercion</span></span><br><span data-ttu-id="15618-281">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="15618-281">
        -PdfFile</span></span><br><span data-ttu-id="15618-282">
        - Sélection</span><span class="sxs-lookup"><span data-stu-id="15618-282">
        - Selection</span></span><br><span data-ttu-id="15618-283">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-283">
        - Settings</span></span><br><span data-ttu-id="15618-284">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="15618-284">
        -TableBindings</span></span><br><span data-ttu-id="15618-285">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-285">
        -TableCoercion</span></span><br><span data-ttu-id="15618-286">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="15618-286">
        -TextBindings</span></span><br><span data-ttu-id="15618-287">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-287">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="15618-288">Outlook</span><span class="sxs-lookup"><span data-stu-id="15618-288">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="15618-289">Plateforme</span><span class="sxs-lookup"><span data-stu-id="15618-289">Platform</span></span></th>
    <th><span data-ttu-id="15618-290">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="15618-290">Extension points</span></span></th>
    <th><span data-ttu-id="15618-291">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="15618-291">API requirement sets</span></span></th>
    <th><span data-ttu-id="15618-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="15618-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-293">Office Online</span><span class="sxs-lookup"><span data-stu-id="15618-293">Office Online</span></span></td>
    <td> <span data-ttu-id="15618-294">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="15618-294">- Mail Read</span></span><br><span data-ttu-id="15618-295">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="15618-295">
      - Mail Compose</span></span><br><span data-ttu-id="15618-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="15618-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="15618-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="15618-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="15618-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="15618-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="15618-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="15618-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="15618-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="15618-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="15618-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="15618-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="15618-304">Non disponible</span><span class="sxs-lookup"><span data-stu-id="15618-304">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-305">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="15618-305">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="15618-306">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="15618-306">- Mail Read</span></span><br><span data-ttu-id="15618-307">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="15618-307">
      - Mail Compose</span></span><br><span data-ttu-id="15618-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="15618-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="15618-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="15618-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="15618-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="15618-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="15618-313">Non disponible</span><span class="sxs-lookup"><span data-stu-id="15618-313">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-314">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="15618-314">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="15618-315">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="15618-315">- Mail Read</span></span><br><span data-ttu-id="15618-316">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="15618-316">
      - Mail Compose</span></span><br><span data-ttu-id="15618-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="15618-318">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="15618-318">
      - Modules</span></span></td>
    <td> <span data-ttu-id="15618-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="15618-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="15618-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="15618-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="15618-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="15618-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="15618-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="15618-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="15618-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="15618-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="15618-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="15618-326">Non disponible</span><span class="sxs-lookup"><span data-stu-id="15618-326">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-327">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="15618-327">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="15618-328">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="15618-328">- Mail Read</span></span><br><span data-ttu-id="15618-329">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="15618-329">
      - Mail Compose</span></span><br><span data-ttu-id="15618-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="15618-331">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="15618-331">
      - Modules</span></span></td>
    <td> <span data-ttu-id="15618-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="15618-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="15618-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="15618-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="15618-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="15618-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="15618-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="15618-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="15618-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="15618-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="15618-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="15618-339">Non disponible</span><span class="sxs-lookup"><span data-stu-id="15618-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-340">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="15618-340">Office for iOS</span></span></td>
    <td> <span data-ttu-id="15618-341">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="15618-341">- Mail Read</span></span><br><span data-ttu-id="15618-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="15618-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="15618-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="15618-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="15618-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="15618-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="15618-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="15618-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="15618-348">Non disponible</span><span class="sxs-lookup"><span data-stu-id="15618-348">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-349">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="15618-349">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="15618-350">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="15618-350">- Mail Read</span></span><br><span data-ttu-id="15618-351">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="15618-351">
      - Mail Compose</span></span><br><span data-ttu-id="15618-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="15618-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="15618-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="15618-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="15618-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="15618-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="15618-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="15618-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="15618-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="15618-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="15618-359">Non disponible</span><span class="sxs-lookup"><span data-stu-id="15618-359">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-360">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="15618-360">Office for Mac</span></span></td>
    <td> <span data-ttu-id="15618-361">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="15618-361">- Mail Read</span></span><br><span data-ttu-id="15618-362">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="15618-362">
      - Mail Compose</span></span><br><span data-ttu-id="15618-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="15618-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="15618-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="15618-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="15618-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="15618-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="15618-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="15618-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="15618-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="15618-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="15618-370">Non disponible</span><span class="sxs-lookup"><span data-stu-id="15618-370">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-371">Office pour Android</span><span class="sxs-lookup"><span data-stu-id="15618-371">Office for Android</span></span></td>
    <td> <span data-ttu-id="15618-372">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="15618-372">- Mail Read</span></span><br><span data-ttu-id="15618-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="15618-374">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-374">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="15618-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="15618-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="15618-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="15618-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="15618-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="15618-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="15618-379">Non disponible</span><span class="sxs-lookup"><span data-stu-id="15618-379">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="15618-380">Word</span><span class="sxs-lookup"><span data-stu-id="15618-380">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="15618-381">Plateforme</span><span class="sxs-lookup"><span data-stu-id="15618-381">Platform</span></span></th>
    <th><span data-ttu-id="15618-382">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="15618-382">Extension points</span></span></th>
    <th><span data-ttu-id="15618-383">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="15618-383">API requirement sets</span></span></th>
    <th><span data-ttu-id="15618-384"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="15618-384"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="15618-385">Office Online</span><span class="sxs-lookup"><span data-stu-id="15618-385">Office Online</span></span></td>
    <td> <span data-ttu-id="15618-386">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-386">- Taskpane</span></span><br><span data-ttu-id="15618-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="15618-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="15618-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="15618-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="15618-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="15618-392">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="15618-392">-BindingEvents</span></span><br><span data-ttu-id="15618-393">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="15618-393">
         -</span></span><br><span data-ttu-id="15618-394">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-394">
         -DocumentEvents</span></span><br><span data-ttu-id="15618-395">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="15618-395">
         - File</span></span><br><span data-ttu-id="15618-396">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-396">
         -HtmlCoercion</span></span><br><span data-ttu-id="15618-397">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-397">
         -ImageCoercion</span></span><br><span data-ttu-id="15618-398">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="15618-398">
         -MatrixBindings</span></span><br><span data-ttu-id="15618-399">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-399">
         -MatrixCoercion</span></span><br><span data-ttu-id="15618-400">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-400">
         -OoxmlCoercion</span></span><br><span data-ttu-id="15618-401">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="15618-401">
         -PdfFile</span></span><br><span data-ttu-id="15618-402">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="15618-402">
         - Selection</span></span><br><span data-ttu-id="15618-403">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-403">
         - Settings</span></span><br><span data-ttu-id="15618-404">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="15618-404">
         -TableBindings</span></span><br><span data-ttu-id="15618-405">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-405">
         -TableCoercion</span></span><br><span data-ttu-id="15618-406">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="15618-406">
         -TextBindings</span></span><br><span data-ttu-id="15618-407">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-407">
         -TextCoercion</span></span><br><span data-ttu-id="15618-408">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="15618-408">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-409">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="15618-409">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="15618-410">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-410">- Taskpane</span></span></td>
    <td> <span data-ttu-id="15618-411">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-411">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="15618-412">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="15618-412">-BindingEvents</span></span><br><span data-ttu-id="15618-413">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-413">
         -CompressedFile</span></span><br><span data-ttu-id="15618-414">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="15618-414">
         -</span></span><br><span data-ttu-id="15618-415">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-415">
         -DocumentEvents</span></span><br><span data-ttu-id="15618-416">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="15618-416">
         - File</span></span><br><span data-ttu-id="15618-417">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-417">
         -HtmlCoercion</span></span><br><span data-ttu-id="15618-418">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-418">
         -ImageCoercion</span></span><br><span data-ttu-id="15618-419">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="15618-419">
         -MatrixBindings</span></span><br><span data-ttu-id="15618-420">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-420">
         -MatrixCoercion</span></span><br><span data-ttu-id="15618-421">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-421">
         -OoxmlCoercion</span></span><br><span data-ttu-id="15618-422">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="15618-422">
         -PdfFile</span></span><br><span data-ttu-id="15618-423">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="15618-423">
         - Selection</span></span><br><span data-ttu-id="15618-424">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-424">
         - Settings</span></span><br><span data-ttu-id="15618-425">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="15618-425">
         -TableBindings</span></span><br><span data-ttu-id="15618-426">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-426">
         -TableCoercion</span></span><br><span data-ttu-id="15618-427">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="15618-427">
         -TextBindings</span></span><br><span data-ttu-id="15618-428">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-428">
         -TextCoercion</span></span><br><span data-ttu-id="15618-429">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="15618-429">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-430">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="15618-430">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="15618-431">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-431">- Taskpane</span></span><br><span data-ttu-id="15618-432">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-432">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="15618-433">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-433">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="15618-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="15618-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="15618-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="15618-437">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="15618-437">-BindingEvents</span></span><br><span data-ttu-id="15618-438">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-438">
         -CompressedFile</span></span><br><span data-ttu-id="15618-439">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="15618-439">
         -</span></span><br><span data-ttu-id="15618-440">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-440">
         -DocumentEvents</span></span><br><span data-ttu-id="15618-441">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="15618-441">
         - File</span></span><br><span data-ttu-id="15618-442">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-442">
         -HtmlCoercion</span></span><br><span data-ttu-id="15618-443">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-443">
         -ImageCoercion</span></span><br><span data-ttu-id="15618-444">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="15618-444">
         -MatrixBindings</span></span><br><span data-ttu-id="15618-445">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-445">
         -MatrixCoercion</span></span><br><span data-ttu-id="15618-446">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-446">
         -OoxmlCoercion</span></span><br><span data-ttu-id="15618-447">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="15618-447">
         -PdfFile</span></span><br><span data-ttu-id="15618-448">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="15618-448">
         - Selection</span></span><br><span data-ttu-id="15618-449">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-449">
         - Settings</span></span><br><span data-ttu-id="15618-450">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="15618-450">
         -TableBindings</span></span><br><span data-ttu-id="15618-451">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-451">
         -TableCoercion</span></span><br><span data-ttu-id="15618-452">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="15618-452">
         -TextBindings</span></span><br><span data-ttu-id="15618-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-453">
         -TextCoercion</span></span><br><span data-ttu-id="15618-454">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="15618-454">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-455">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="15618-455">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="15618-456">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-456">- Taskpane</span></span><br><span data-ttu-id="15618-457">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-457">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="15618-458">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-458">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="15618-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="15618-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="15618-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="15618-462">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="15618-462">-BindingEvents</span></span><br><span data-ttu-id="15618-463">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-463">
         -CompressedFile</span></span><br><span data-ttu-id="15618-464">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="15618-464">
         -</span></span><br><span data-ttu-id="15618-465">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-465">
         -DocumentEvents</span></span><br><span data-ttu-id="15618-466">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="15618-466">
         - File</span></span><br><span data-ttu-id="15618-467">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-467">
         -HtmlCoercion</span></span><br><span data-ttu-id="15618-468">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-468">
         -ImageCoercion</span></span><br><span data-ttu-id="15618-469">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="15618-469">
         -MatrixBindings</span></span><br><span data-ttu-id="15618-470">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-470">
         -MatrixCoercion</span></span><br><span data-ttu-id="15618-471">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-471">
         -OoxmlCoercion</span></span><br><span data-ttu-id="15618-472">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="15618-472">
         -PdfFile</span></span><br><span data-ttu-id="15618-473">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="15618-473">
         - Selection</span></span><br><span data-ttu-id="15618-474">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-474">
         - Settings</span></span><br><span data-ttu-id="15618-475">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="15618-475">
         -TableBindings</span></span><br><span data-ttu-id="15618-476">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-476">
         -TableCoercion</span></span><br><span data-ttu-id="15618-477">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="15618-477">
         -TextBindings</span></span><br><span data-ttu-id="15618-478">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-478">
         -TextCoercion</span></span><br><span data-ttu-id="15618-479">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="15618-479">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-480">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="15618-480">Office for iOS</span></span></td>
    <td> <span data-ttu-id="15618-481">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-481">- Taskpane</span></span></td>
    <td> <span data-ttu-id="15618-482">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-482">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="15618-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="15618-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="15618-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="15618-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="15618-486">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="15618-486">-BindingEvents</span></span><br><span data-ttu-id="15618-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-487">
         -CompressedFile</span></span><br><span data-ttu-id="15618-488">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="15618-488">
         -</span></span><br><span data-ttu-id="15618-489">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-489">
         -DocumentEvents</span></span><br><span data-ttu-id="15618-490">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="15618-490">
         - File</span></span><br><span data-ttu-id="15618-491">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-491">
         -HtmlCoercion</span></span><br><span data-ttu-id="15618-492">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-492">
         -ImageCoercion</span></span><br><span data-ttu-id="15618-493">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="15618-493">
         -MatrixBindings</span></span><br><span data-ttu-id="15618-494">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-494">
         -MatrixCoercion</span></span><br><span data-ttu-id="15618-495">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-495">
         -OoxmlCoercion</span></span><br><span data-ttu-id="15618-496">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="15618-496">
         -PdfFile</span></span><br><span data-ttu-id="15618-497">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="15618-497">
         - Selection</span></span><br><span data-ttu-id="15618-498">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-498">
         - Settings</span></span><br><span data-ttu-id="15618-499">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="15618-499">
         -TableBindings</span></span><br><span data-ttu-id="15618-500">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-500">
         -TableCoercion</span></span><br><span data-ttu-id="15618-501">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="15618-501">
         -TextBindings</span></span><br><span data-ttu-id="15618-502">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-502">
         -TextCoercion</span></span><br><span data-ttu-id="15618-503">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="15618-503">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-504">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="15618-504">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="15618-505">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-505">- Taskpane</span></span><br><span data-ttu-id="15618-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="15618-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="15618-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="15618-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="15618-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="15618-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="15618-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="15618-511">-BindingEvents</span></span><br><span data-ttu-id="15618-512">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-512">
         -CompressedFile</span></span><br><span data-ttu-id="15618-513">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="15618-513">
         -</span></span><br><span data-ttu-id="15618-514">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-514">
         -DocumentEvents</span></span><br><span data-ttu-id="15618-515">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="15618-515">
         - File</span></span><br><span data-ttu-id="15618-516">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-516">
         -HtmlCoercion</span></span><br><span data-ttu-id="15618-517">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-517">
         -ImageCoercion</span></span><br><span data-ttu-id="15618-518">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="15618-518">
         -MatrixBindings</span></span><br><span data-ttu-id="15618-519">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-519">
         -MatrixCoercion</span></span><br><span data-ttu-id="15618-520">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-520">
         -OoxmlCoercion</span></span><br><span data-ttu-id="15618-521">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="15618-521">
         -PdfFile</span></span><br><span data-ttu-id="15618-522">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="15618-522">
         - Selection</span></span><br><span data-ttu-id="15618-523">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-523">
         - Settings</span></span><br><span data-ttu-id="15618-524">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="15618-524">
         -TableBindings</span></span><br><span data-ttu-id="15618-525">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-525">
         -TableCoercion</span></span><br><span data-ttu-id="15618-526">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="15618-526">
         -TextBindings</span></span><br><span data-ttu-id="15618-527">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-527">
         -TextCoercion</span></span><br><span data-ttu-id="15618-528">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="15618-528">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-529">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="15618-529">Office for Mac</span></span></td>
    <td> <span data-ttu-id="15618-530">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-530">- Taskpane</span></span><br><span data-ttu-id="15618-531">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-531">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="15618-532">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-532">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="15618-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="15618-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="15618-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="15618-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="15618-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="15618-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="15618-536">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="15618-536">-BindingEvents</span></span><br><span data-ttu-id="15618-537">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-537">
         -CompressedFile</span></span><br><span data-ttu-id="15618-538">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="15618-538">
         -</span></span><br><span data-ttu-id="15618-539">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-539">
         -DocumentEvents</span></span><br><span data-ttu-id="15618-540">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="15618-540">
         - File</span></span><br><span data-ttu-id="15618-541">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-541">
         -HtmlCoercion</span></span><br><span data-ttu-id="15618-542">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-542">
         -ImageCoercion</span></span><br><span data-ttu-id="15618-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="15618-543">
         -MatrixBindings</span></span><br><span data-ttu-id="15618-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-544">
         -MatrixCoercion</span></span><br><span data-ttu-id="15618-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-545">
         -OoxmlCoercion</span></span><br><span data-ttu-id="15618-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="15618-546">
         -PdfFile</span></span><br><span data-ttu-id="15618-547">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="15618-547">
         - Selection</span></span><br><span data-ttu-id="15618-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-548">
         - Settings</span></span><br><span data-ttu-id="15618-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="15618-549">
         -TableBindings</span></span><br><span data-ttu-id="15618-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-550">
         -TableCoercion</span></span><br><span data-ttu-id="15618-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="15618-551">
         -TextBindings</span></span><br><span data-ttu-id="15618-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-552">
         -TextCoercion</span></span><br><span data-ttu-id="15618-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="15618-553">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="15618-554">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="15618-554">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="15618-555">Plateforme</span><span class="sxs-lookup"><span data-stu-id="15618-555">Platform</span></span></th>
    <th><span data-ttu-id="15618-556">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="15618-556">Extension points</span></span></th>
    <th><span data-ttu-id="15618-557">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="15618-557">API requirement sets</span></span></th>
    <th><span data-ttu-id="15618-558"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="15618-558"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="15618-559">Office Online</span><span class="sxs-lookup"><span data-stu-id="15618-559">Office Online</span></span></td>
    <td> <span data-ttu-id="15618-560">- Contenu</span><span class="sxs-lookup"><span data-stu-id="15618-560">- Content</span></span><br><span data-ttu-id="15618-561">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-561">
         - Taskpane</span></span><br><span data-ttu-id="15618-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="15618-563">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-563">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="15618-564">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="15618-564">-ActiveView</span></span><br><span data-ttu-id="15618-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-565">
         -CompressedFile</span></span><br><span data-ttu-id="15618-566">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-566">
         -DocumentEvents</span></span><br><span data-ttu-id="15618-567">
         - File</span><span class="sxs-lookup"><span data-stu-id="15618-567">
         - File</span></span><br><span data-ttu-id="15618-568">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-568">
         -ImageCoercion</span></span><br><span data-ttu-id="15618-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="15618-569">
         -PdfFile</span></span><br><span data-ttu-id="15618-570">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="15618-570">
         - Selection</span></span><br><span data-ttu-id="15618-571">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-571">
         - Settings</span></span><br><span data-ttu-id="15618-572">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-572">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-573">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="15618-573">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="15618-574">- Contenu</span><span class="sxs-lookup"><span data-stu-id="15618-574">- Content</span></span><br><span data-ttu-id="15618-575">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-575">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="15618-576">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="15618-576">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="15618-577">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="15618-577">-ActiveView</span></span><br><span data-ttu-id="15618-578">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-578">
         -CompressedFile</span></span><br><span data-ttu-id="15618-579">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-579">
         -DocumentEvents</span></span><br><span data-ttu-id="15618-580">
         - File</span><span class="sxs-lookup"><span data-stu-id="15618-580">
         - File</span></span><br><span data-ttu-id="15618-581">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-581">
         -ImageCoercion</span></span><br><span data-ttu-id="15618-582">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="15618-582">
         -PdfFile</span></span><br><span data-ttu-id="15618-583">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="15618-583">
         - Selection</span></span><br><span data-ttu-id="15618-584">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-584">
         - Settings</span></span><br><span data-ttu-id="15618-585">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-585">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-586">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="15618-586">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="15618-587">- Contenu</span><span class="sxs-lookup"><span data-stu-id="15618-587">- Content</span></span><br><span data-ttu-id="15618-588">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-588">
         - Taskpane</span></span><br><span data-ttu-id="15618-589">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-589">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="15618-590">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-590">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="15618-591">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="15618-591">-ActiveView</span></span><br><span data-ttu-id="15618-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-592">
         -CompressedFile</span></span><br><span data-ttu-id="15618-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-593">
         -DocumentEvents</span></span><br><span data-ttu-id="15618-594">
         - File</span><span class="sxs-lookup"><span data-stu-id="15618-594">
         - File</span></span><br><span data-ttu-id="15618-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-595">
         -ImageCoercion</span></span><br><span data-ttu-id="15618-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="15618-596">
         -PdfFile</span></span><br><span data-ttu-id="15618-597">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="15618-597">
         - Selection</span></span><br><span data-ttu-id="15618-598">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-598">
         - Settings</span></span><br><span data-ttu-id="15618-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-599">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-600">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="15618-600">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="15618-601">- Contenu</span><span class="sxs-lookup"><span data-stu-id="15618-601">- Content</span></span><br><span data-ttu-id="15618-602">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-602">
         - Taskpane</span></span><br><span data-ttu-id="15618-603">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-603">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="15618-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="15618-605">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="15618-605">-ActiveView</span></span><br><span data-ttu-id="15618-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-606">
         -CompressedFile</span></span><br><span data-ttu-id="15618-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-607">
         -DocumentEvents</span></span><br><span data-ttu-id="15618-608">
         - File</span><span class="sxs-lookup"><span data-stu-id="15618-608">
         - File</span></span><br><span data-ttu-id="15618-609">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-609">
         -ImageCoercion</span></span><br><span data-ttu-id="15618-610">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="15618-610">
         -PdfFile</span></span><br><span data-ttu-id="15618-611">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="15618-611">
         - Selection</span></span><br><span data-ttu-id="15618-612">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-612">
         - Settings</span></span><br><span data-ttu-id="15618-613">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-613">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-614">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="15618-614">Office for iOS</span></span></td>
    <td> <span data-ttu-id="15618-615">- Contenu</span><span class="sxs-lookup"><span data-stu-id="15618-615">- Content</span></span><br><span data-ttu-id="15618-616">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-616">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="15618-617">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-617">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="15618-618">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="15618-618">-ActiveView</span></span><br><span data-ttu-id="15618-619">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-619">
         -CompressedFile</span></span><br><span data-ttu-id="15618-620">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-620">
         -DocumentEvents</span></span><br><span data-ttu-id="15618-621">
         - Fichier</span><span class="sxs-lookup"><span data-stu-id="15618-621">
         - File</span></span><br><span data-ttu-id="15618-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="15618-622">
         -PdfFile</span></span><br><span data-ttu-id="15618-623">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="15618-623">
         - Selection</span></span><br><span data-ttu-id="15618-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-624">
         - Settings</span></span><br><span data-ttu-id="15618-625">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-625">
         -TextCoercion</span></span><br><span data-ttu-id="15618-626">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-626">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-627">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="15618-627">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="15618-628">- Contenu</span><span class="sxs-lookup"><span data-stu-id="15618-628">- Content</span></span><br><span data-ttu-id="15618-629">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-629">
         - Taskpane</span></span><br><span data-ttu-id="15618-630">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-630">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="15618-631">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-631">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="15618-632">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="15618-632">-ActiveView</span></span><br><span data-ttu-id="15618-633">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-633">
         -CompressedFile</span></span><br><span data-ttu-id="15618-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-634">
         -DocumentEvents</span></span><br><span data-ttu-id="15618-635">
         - File</span><span class="sxs-lookup"><span data-stu-id="15618-635">
         - File</span></span><br><span data-ttu-id="15618-636">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-636">
         -ImageCoercion</span></span><br><span data-ttu-id="15618-637">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="15618-637">
         -PdfFile</span></span><br><span data-ttu-id="15618-638">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="15618-638">
         - Selection</span></span><br><span data-ttu-id="15618-639">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-639">
         - Settings</span></span><br><span data-ttu-id="15618-640">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-640">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="15618-641">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="15618-641">Office for Mac</span></span></td>
    <td> <span data-ttu-id="15618-642">- Contenu</span><span class="sxs-lookup"><span data-stu-id="15618-642">- Content</span></span><br><span data-ttu-id="15618-643">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-643">
         - Taskpane</span></span><br><span data-ttu-id="15618-644">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-644">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="15618-645">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-645">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="15618-646">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="15618-646">-ActiveView</span></span><br><span data-ttu-id="15618-647">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="15618-647">
         -CompressedFile</span></span><br><span data-ttu-id="15618-648">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-648">
         -DocumentEvents</span></span><br><span data-ttu-id="15618-649">
         - File</span><span class="sxs-lookup"><span data-stu-id="15618-649">
         - File</span></span><br><span data-ttu-id="15618-650">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-650">
         -ImageCoercion</span></span><br><span data-ttu-id="15618-651">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="15618-651">
         -PdfFile</span></span><br><span data-ttu-id="15618-652">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="15618-652">
         - Selection</span></span><br><span data-ttu-id="15618-653">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-653">
         - Settings</span></span><br><span data-ttu-id="15618-654">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-654">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="15618-655">OneNote</span><span class="sxs-lookup"><span data-stu-id="15618-655">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="15618-656">Plateforme</span><span class="sxs-lookup"><span data-stu-id="15618-656">Platform</span></span></th>
    <th><span data-ttu-id="15618-657">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="15618-657">Extension points</span></span></th>
    <th><span data-ttu-id="15618-658">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="15618-658">API requirement sets</span></span></th>
    <th><span data-ttu-id="15618-659"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="15618-659"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="15618-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="15618-660">Office Online</span></span></td>
    <td> <span data-ttu-id="15618-661">- Contenu</span><span class="sxs-lookup"><span data-stu-id="15618-661">- Content</span></span><br><span data-ttu-id="15618-662">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="15618-662">
         - Taskpane</span></span><br><span data-ttu-id="15618-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="15618-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="15618-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="15618-665">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="15618-665">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="15618-666">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="15618-666">-DocumentEvents</span></span><br><span data-ttu-id="15618-667">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-667">
         -HtmlCoercion</span></span><br><span data-ttu-id="15618-668">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-668">
         -ImageCoercion</span></span><br><span data-ttu-id="15618-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="15618-669">
         - Settings</span></span><br><span data-ttu-id="15618-670">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="15618-670">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="15618-671">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="15618-671">See also</span></span>

- [<span data-ttu-id="15618-672">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="15618-672">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="15618-673">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="15618-673">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="15618-674">Ensembles de conditions requises des commandes de complément</span><span class="sxs-lookup"><span data-stu-id="15618-674">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="15618-675">Référence de l’interface API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="15618-675">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
