---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint et OneNote.
ms.date: 09/19/2018
ms.openlocfilehash: 09fb72c88bd0496c413f94b7ba4149192380d664
ms.sourcegitcommit: e7e4d08569a01c69168bb005188e9a1e628304b9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/22/2018
ms.locfileid: "24967703"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="1c5bf-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="1c5bf-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="1c5bf-104">Pour fonctionner comme prévu, il se peut que votre complément Office dépende d’un hôte Office spécifique, d’un ensemble de conditions requises, d’un membre d’API ou d’une version de l’API.</span><span class="sxs-lookup"><span data-stu-id="1c5bf-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="1c5bf-105">Les tableaux suivants contiennent la plateforme disponible, les points d’extension, les ensembles de conditions requises de l’API et les ensembles de conditions requises des API communes qui sont actuellement pris en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="1c5bf-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="1c5bf-106">Si une cellule de tableau contient un astérisque (\*), cela signifie que nous travaillons sur celle-ci.</span><span class="sxs-lookup"><span data-stu-id="1c5bf-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="1c5bf-107">Pour les ensembles de conditions requises pour Project ou Access, consultez les [ensembles de conditions requises communs à Office](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="1c5bf-107">For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="1c5bf-p103">Le numéro de build pour Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que les ensembles de conditions requises des API communes, ExcelApi 1.1 et WordApi 1.1.</span><span class="sxs-lookup"><span data-stu-id="1c5bf-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="1c5bf-110">Excel</span><span class="sxs-lookup"><span data-stu-id="1c5bf-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="1c5bf-111">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="1c5bf-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="1c5bf-112">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="1c5bf-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="1c5bf-113">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="1c5bf-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="1c5bf-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="1c5bf-115">Office Online</span></span></td>
    <td> <span data-ttu-id="1c5bf-116">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-116">- Taskpane</span></span><br><span data-ttu-id="1c5bf-117">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="1c5bf-117">
        - Content</span></span><br><span data-ttu-id="1c5bf-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="1c5bf-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="1c5bf-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1c5bf-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1c5bf-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1c5bf-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1c5bf-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1c5bf-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1c5bf-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="1c5bf-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1c5bf-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-127">
        -BindingEvents</span></span><br><span data-ttu-id="1c5bf-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-128">
        -CompressedFile</span></span><br><span data-ttu-id="1c5bf-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-129">
        -DocumentEvents</span></span><br><span data-ttu-id="1c5bf-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="1c5bf-130">
        - File</span></span><br><span data-ttu-id="1c5bf-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-131">
        -MatrixBindings</span></span><br><span data-ttu-id="1c5bf-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-132">
        -MatrixCoercion</span></span><br><span data-ttu-id="1c5bf-133">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1c5bf-133">
        - Selection</span></span><br><span data-ttu-id="1c5bf-134">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-134">
        - Settings</span></span><br><span data-ttu-id="1c5bf-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-135">
        -TableBindings</span></span><br><span data-ttu-id="1c5bf-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-136">
        -TableCoercion</span></span><br><span data-ttu-id="1c5bf-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-137">
        -TextBindings</span></span><br><span data-ttu-id="1c5bf-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-138">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-139">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1c5bf-139">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="1c5bf-140">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-140">
        - Taskpane</span></span><br><span data-ttu-id="1c5bf-141">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="1c5bf-141">
        - Content</span></span></td>
    <td>  <span data-ttu-id="1c5bf-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1c5bf-143">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-143">
        -BindingEvents</span></span><br><span data-ttu-id="1c5bf-144">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-144">
        -CompressedFile</span></span><br><span data-ttu-id="1c5bf-145">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-145">
        -DocumentEvents</span></span><br><span data-ttu-id="1c5bf-146">
        - File</span><span class="sxs-lookup"><span data-stu-id="1c5bf-146">
        - File</span></span><br><span data-ttu-id="1c5bf-147">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-147">
        -ImageCoercion</span></span><br><span data-ttu-id="1c5bf-148">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-148">
        -MatrixBindings</span></span><br><span data-ttu-id="1c5bf-149">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-149">
        -MatrixCoercion</span></span><br><span data-ttu-id="1c5bf-150">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1c5bf-150">
        - Selection</span></span><br><span data-ttu-id="1c5bf-151">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-151">
        - Settings</span></span><br><span data-ttu-id="1c5bf-152">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-152">
        -TableBindings</span></span><br><span data-ttu-id="1c5bf-153">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-153">
        -TableCoercion</span></span><br><span data-ttu-id="1c5bf-154">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-154">
        -TextBindings</span></span><br><span data-ttu-id="1c5bf-155">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-155">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-156">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1c5bf-156">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="1c5bf-157">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-157">- Taskpane</span></span><br><span data-ttu-id="1c5bf-158">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="1c5bf-158">
        - Content</span></span><br><span data-ttu-id="1c5bf-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1c5bf-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1c5bf-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1c5bf-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1c5bf-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1c5bf-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1c5bf-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1c5bf-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-166">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="1c5bf-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1c5bf-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-168">-BindingEvents</span></span><br><span data-ttu-id="1c5bf-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-169">
        -CompressedFile</span></span><br><span data-ttu-id="1c5bf-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-170">
        -DocumentEvents</span></span><br><span data-ttu-id="1c5bf-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="1c5bf-171">
        - File</span></span><br><span data-ttu-id="1c5bf-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-172">
        -ImageCoercion</span></span><br><span data-ttu-id="1c5bf-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-173">
        -MatrixBindings</span></span><br><span data-ttu-id="1c5bf-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="1c5bf-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1c5bf-175">
        - Selection</span></span><br><span data-ttu-id="1c5bf-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-176">
        - Settings</span></span><br><span data-ttu-id="1c5bf-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-177">
        -TableBindings</span></span><br><span data-ttu-id="1c5bf-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-178">
        -TableCoercion</span></span><br><span data-ttu-id="1c5bf-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-179">
        -TextBindings</span></span><br><span data-ttu-id="1c5bf-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-181">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="1c5bf-181">Office for iOS</span></span></td>
    <td><span data-ttu-id="1c5bf-182">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-182">- Taskpane</span></span><br><span data-ttu-id="1c5bf-183">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="1c5bf-183">
        - Content</span></span></td>
    <td><span data-ttu-id="1c5bf-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1c5bf-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1c5bf-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1c5bf-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1c5bf-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1c5bf-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1c5bf-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-190">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="1c5bf-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1c5bf-192">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-192">-BindingEvents</span></span><br><span data-ttu-id="1c5bf-193">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-193">
        -CompressedFile</span></span><br><span data-ttu-id="1c5bf-194">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-194">
        -DocumentEvents</span></span><br><span data-ttu-id="1c5bf-195">
        - File</span><span class="sxs-lookup"><span data-stu-id="1c5bf-195">
        - File</span></span><br><span data-ttu-id="1c5bf-196">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-196">
        -ImageCoercion</span></span><br><span data-ttu-id="1c5bf-197">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-197">
        -MatrixBindings</span></span><br><span data-ttu-id="1c5bf-198">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-198">
        -MatrixCoercion</span></span><br><span data-ttu-id="1c5bf-199">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1c5bf-199">
        - Selection</span></span><br><span data-ttu-id="1c5bf-200">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-200">
        - Settings</span></span><br><span data-ttu-id="1c5bf-201">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-201">
        -TableBindings</span></span><br><span data-ttu-id="1c5bf-202">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-202">
        -TableCoercion</span></span><br><span data-ttu-id="1c5bf-203">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-203">
        -TextBindings</span></span><br><span data-ttu-id="1c5bf-204">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-204">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-205">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="1c5bf-205">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="1c5bf-206">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-206">- Taskpane</span></span><br><span data-ttu-id="1c5bf-207">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="1c5bf-207">
        - Content</span></span><br><span data-ttu-id="1c5bf-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1c5bf-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1c5bf-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1c5bf-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1c5bf-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1c5bf-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1c5bf-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1c5bf-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-215">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="1c5bf-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1c5bf-217">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-217">-BindingEvents</span></span><br><span data-ttu-id="1c5bf-218">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-218">
        -CompressedFile</span></span><br><span data-ttu-id="1c5bf-219">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-219">
        -DocumentEvents</span></span><br><span data-ttu-id="1c5bf-220">
        - File</span><span class="sxs-lookup"><span data-stu-id="1c5bf-220">
        - File</span></span><br><span data-ttu-id="1c5bf-221">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-221">
        -ImageCoercion</span></span><br><span data-ttu-id="1c5bf-222">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-222">
        -MatrixBindings</span></span><br><span data-ttu-id="1c5bf-223">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-223">
        -MatrixCoercion</span></span><br><span data-ttu-id="1c5bf-224">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-224">
        -PdfFile</span></span><br><span data-ttu-id="1c5bf-225">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1c5bf-225">
        - Selection</span></span><br><span data-ttu-id="1c5bf-226">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-226">
        - Settings</span></span><br><span data-ttu-id="1c5bf-227">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-227">
        -TableBindings</span></span><br><span data-ttu-id="1c5bf-228">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-228">
        -TableCoercion</span></span><br><span data-ttu-id="1c5bf-229">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-229">
        -TextBindings</span></span><br><span data-ttu-id="1c5bf-230">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-230">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="1c5bf-231">Outlook</span><span class="sxs-lookup"><span data-stu-id="1c5bf-231">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1c5bf-232">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="1c5bf-232">Platform</span></span></th>
    <th><span data-ttu-id="1c5bf-233">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="1c5bf-233">Extension points</span></span></th>
    <th><span data-ttu-id="1c5bf-234">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="1c5bf-234">API requirement sets</span></span></th>
    <th><span data-ttu-id="1c5bf-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-236">Office Online</span><span class="sxs-lookup"><span data-stu-id="1c5bf-236">Office Online</span></span></td>
    <td> <span data-ttu-id="1c5bf-237">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="1c5bf-237">- Mail Read</span></span><br><span data-ttu-id="1c5bf-238">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="1c5bf-238">
      - Mail Compose</span></span><br><span data-ttu-id="1c5bf-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c5bf-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1c5bf-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1c5bf-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1c5bf-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1c5bf-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1c5bf-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="1c5bf-246">Non disponible</span><span class="sxs-lookup"><span data-stu-id="1c5bf-246">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-247">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1c5bf-247">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="1c5bf-248">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="1c5bf-248">- Mail Read</span></span><br><span data-ttu-id="1c5bf-249">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="1c5bf-249">
      - Mail Compose</span></span><br><span data-ttu-id="1c5bf-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c5bf-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1c5bf-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1c5bf-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1c5bf-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="1c5bf-255">Non disponible</span><span class="sxs-lookup"><span data-stu-id="1c5bf-255">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-256">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1c5bf-256">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="1c5bf-257">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="1c5bf-257">- Mail Read</span></span><br><span data-ttu-id="1c5bf-258">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="1c5bf-258">
      - Mail Compose</span></span><br><span data-ttu-id="1c5bf-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="1c5bf-260">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="1c5bf-260">
      - Modules</span></span></td>
    <td> <span data-ttu-id="1c5bf-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1c5bf-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1c5bf-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1c5bf-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1c5bf-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1c5bf-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1c5bf-267">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-267">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1c5bf-268">Non disponible</span><span class="sxs-lookup"><span data-stu-id="1c5bf-268">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-269">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="1c5bf-269">Office for iOS</span></span></td>
    <td> <span data-ttu-id="1c5bf-270">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="1c5bf-270">- Mail Read</span></span><br><span data-ttu-id="1c5bf-271">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-271">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c5bf-272">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-272">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1c5bf-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1c5bf-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1c5bf-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1c5bf-276">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-276">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="1c5bf-277">Non disponible</span><span class="sxs-lookup"><span data-stu-id="1c5bf-277">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-278">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="1c5bf-278">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="1c5bf-279">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="1c5bf-279">- Mail Read</span></span><br><span data-ttu-id="1c5bf-280">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="1c5bf-280">
      - Mail Compose</span></span><br><span data-ttu-id="1c5bf-281">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-281">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c5bf-282">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-282">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1c5bf-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1c5bf-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1c5bf-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1c5bf-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1c5bf-287">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-287">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="1c5bf-288">Non disponible</span><span class="sxs-lookup"><span data-stu-id="1c5bf-288">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-289">Office pour Android</span><span class="sxs-lookup"><span data-stu-id="1c5bf-289">Office for Android</span></span></td>
    <td> <span data-ttu-id="1c5bf-290">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="1c5bf-290">- Mail Read</span></span><br><span data-ttu-id="1c5bf-291">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-291">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c5bf-292">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-292">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1c5bf-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1c5bf-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1c5bf-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1c5bf-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="1c5bf-297">Non disponible</span><span class="sxs-lookup"><span data-stu-id="1c5bf-297">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="1c5bf-298">Word</span><span class="sxs-lookup"><span data-stu-id="1c5bf-298">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1c5bf-299">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="1c5bf-299">Platform</span></span></th>
    <th><span data-ttu-id="1c5bf-300">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="1c5bf-300">Extension points</span></span></th>
    <th><span data-ttu-id="1c5bf-301">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="1c5bf-301">API requirement sets</span></span></th>
    <th><span data-ttu-id="1c5bf-302"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-302"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-303">Office Online</span><span class="sxs-lookup"><span data-stu-id="1c5bf-303">Office Online</span></span></td>
    <td> <span data-ttu-id="1c5bf-304">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-304">- Taskpane</span></span><br><span data-ttu-id="1c5bf-305">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-305">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c5bf-306">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-306">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1c5bf-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1c5bf-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1c5bf-309">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-309">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c5bf-310">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-310">-BindingEvents</span></span><br><span data-ttu-id="1c5bf-311">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1c5bf-311">
         -</span></span><br><span data-ttu-id="1c5bf-312">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-312">
         -DocumentEvents</span></span><br><span data-ttu-id="1c5bf-313">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c5bf-313">
         - File</span></span><br><span data-ttu-id="1c5bf-314">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-314">
         -HtmlCoercion</span></span><br><span data-ttu-id="1c5bf-315">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-315">
         -ImageCoercion</span></span><br><span data-ttu-id="1c5bf-316">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-316">
         -MatrixBindings</span></span><br><span data-ttu-id="1c5bf-317">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-317">
         -MatrixCoercion</span></span><br><span data-ttu-id="1c5bf-318">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-318">
         -OoxmlCoercion</span></span><br><span data-ttu-id="1c5bf-319">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-319">
         -PdfFile</span></span><br><span data-ttu-id="1c5bf-320">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c5bf-320">
         - Selection</span></span><br><span data-ttu-id="1c5bf-321">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-321">
         - Settings</span></span><br><span data-ttu-id="1c5bf-322">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-322">
         -TableBindings</span></span><br><span data-ttu-id="1c5bf-323">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-323">
         -TableCoercion</span></span><br><span data-ttu-id="1c5bf-324">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-324">
         -TextBindings</span></span><br><span data-ttu-id="1c5bf-325">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-325">
         -TextCoercion</span></span><br><span data-ttu-id="1c5bf-326">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-326">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-327">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1c5bf-327">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="1c5bf-328">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-328">- Taskpane</span></span></td>
    <td> <span data-ttu-id="1c5bf-329">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-329">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c5bf-330">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-330">-BindingEvents</span></span><br><span data-ttu-id="1c5bf-331">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-331">
         -CompressedFile</span></span><br><span data-ttu-id="1c5bf-332">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1c5bf-332">
         -</span></span><br><span data-ttu-id="1c5bf-333">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-333">
         -DocumentEvents</span></span><br><span data-ttu-id="1c5bf-334">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c5bf-334">
         - File</span></span><br><span data-ttu-id="1c5bf-335">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-335">
         -HtmlCoercion</span></span><br><span data-ttu-id="1c5bf-336">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-336">
         -ImageCoercion</span></span><br><span data-ttu-id="1c5bf-337">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-337">
         -MatrixBindings</span></span><br><span data-ttu-id="1c5bf-338">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-338">
         -MatrixCoercion</span></span><br><span data-ttu-id="1c5bf-339">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-339">
         -OoxmlCoercion</span></span><br><span data-ttu-id="1c5bf-340">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-340">
         -PdfFile</span></span><br><span data-ttu-id="1c5bf-341">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c5bf-341">
         - Selection</span></span><br><span data-ttu-id="1c5bf-342">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-342">
         - Settings</span></span><br><span data-ttu-id="1c5bf-343">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-343">
         -TableBindings</span></span><br><span data-ttu-id="1c5bf-344">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-344">
         -TableCoercion</span></span><br><span data-ttu-id="1c5bf-345">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-345">
         -TextBindings</span></span><br><span data-ttu-id="1c5bf-346">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-346">
         -TextCoercion</span></span><br><span data-ttu-id="1c5bf-347">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-347">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-348">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1c5bf-348">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="1c5bf-349">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-349">- Taskpane</span></span><br><span data-ttu-id="1c5bf-350">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-350">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c5bf-351">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-351">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1c5bf-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1c5bf-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1c5bf-354">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-354">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c5bf-355">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-355">-BindingEvents</span></span><br><span data-ttu-id="1c5bf-356">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-356">
         -CompressedFile</span></span><br><span data-ttu-id="1c5bf-357">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1c5bf-357">
         -</span></span><br><span data-ttu-id="1c5bf-358">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-358">
         -DocumentEvents</span></span><br><span data-ttu-id="1c5bf-359">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c5bf-359">
         - File</span></span><br><span data-ttu-id="1c5bf-360">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-360">
         -HtmlCoercion</span></span><br><span data-ttu-id="1c5bf-361">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-361">
         -ImageCoercion</span></span><br><span data-ttu-id="1c5bf-362">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-362">
         -MatrixBindings</span></span><br><span data-ttu-id="1c5bf-363">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-363">
         -MatrixCoercion</span></span><br><span data-ttu-id="1c5bf-364">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-364">
         -OoxmlCoercion</span></span><br><span data-ttu-id="1c5bf-365">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-365">
         -PdfFile</span></span><br><span data-ttu-id="1c5bf-366">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c5bf-366">
         - Selection</span></span><br><span data-ttu-id="1c5bf-367">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-367">
         - Settings</span></span><br><span data-ttu-id="1c5bf-368">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-368">
         -TableBindings</span></span><br><span data-ttu-id="1c5bf-369">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-369">
         -TableCoercion</span></span><br><span data-ttu-id="1c5bf-370">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-370">
         -TextBindings</span></span><br><span data-ttu-id="1c5bf-371">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-371">
         -TextCoercion</span></span><br><span data-ttu-id="1c5bf-372">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-372">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-373">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="1c5bf-373">Office for iOS</span></span></td>
    <td> <span data-ttu-id="1c5bf-374">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-374">- Taskpane</span></span></td>
    <td> <span data-ttu-id="1c5bf-375">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-375">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1c5bf-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1c5bf-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1c5bf-378">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="1c5bf-378">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="1c5bf-379">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-379">-BindingEvents</span></span><br><span data-ttu-id="1c5bf-380">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-380">
         -CompressedFile</span></span><br><span data-ttu-id="1c5bf-381">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1c5bf-381">
         -</span></span><br><span data-ttu-id="1c5bf-382">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-382">
         -DocumentEvents</span></span><br><span data-ttu-id="1c5bf-383">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c5bf-383">
         - File</span></span><br><span data-ttu-id="1c5bf-384">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-384">
         -HtmlCoercion</span></span><br><span data-ttu-id="1c5bf-385">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-385">
         -ImageCoercion</span></span><br><span data-ttu-id="1c5bf-386">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-386">
         -MatrixBindings</span></span><br><span data-ttu-id="1c5bf-387">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-387">
         -MatrixCoercion</span></span><br><span data-ttu-id="1c5bf-388">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-388">
         -OoxmlCoercion</span></span><br><span data-ttu-id="1c5bf-389">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-389">
         -PdfFile</span></span><br><span data-ttu-id="1c5bf-390">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c5bf-390">
         - Selection</span></span><br><span data-ttu-id="1c5bf-391">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-391">
         - Settings</span></span><br><span data-ttu-id="1c5bf-392">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-392">
         -TableBindings</span></span><br><span data-ttu-id="1c5bf-393">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-393">
         -TableCoercion</span></span><br><span data-ttu-id="1c5bf-394">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-394">
         -TextBindings</span></span><br><span data-ttu-id="1c5bf-395">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-395">
         -TextCoercion</span></span><br><span data-ttu-id="1c5bf-396">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-396">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-397">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="1c5bf-397">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="1c5bf-398">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-398">- Taskpane</span></span><br><span data-ttu-id="1c5bf-399">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-399">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c5bf-400">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-400">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1c5bf-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1c5bf-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1c5bf-403">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="1c5bf-403">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="1c5bf-404">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-404">-BindingEvents</span></span><br><span data-ttu-id="1c5bf-405">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-405">
         -CompressedFile</span></span><br><span data-ttu-id="1c5bf-406">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1c5bf-406">
         -</span></span><br><span data-ttu-id="1c5bf-407">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-407">
         -DocumentEvents</span></span><br><span data-ttu-id="1c5bf-408">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c5bf-408">
         - File</span></span><br><span data-ttu-id="1c5bf-409">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-409">
         -HtmlCoercion</span></span><br><span data-ttu-id="1c5bf-410">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-410">
         -ImageCoercion</span></span><br><span data-ttu-id="1c5bf-411">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-411">
         -MatrixBindings</span></span><br><span data-ttu-id="1c5bf-412">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-412">
         -MatrixCoercion</span></span><br><span data-ttu-id="1c5bf-413">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-413">
         -OoxmlCoercion</span></span><br><span data-ttu-id="1c5bf-414">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-414">
         -PdfFile</span></span><br><span data-ttu-id="1c5bf-415">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c5bf-415">
         - Selection</span></span><br><span data-ttu-id="1c5bf-416">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-416">
         - Settings</span></span><br><span data-ttu-id="1c5bf-417">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-417">
         -TableBindings</span></span><br><span data-ttu-id="1c5bf-418">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-418">
         -TableCoercion</span></span><br><span data-ttu-id="1c5bf-419">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-419">
         -TextBindings</span></span><br><span data-ttu-id="1c5bf-420">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-420">
         -TextCoercion</span></span><br><span data-ttu-id="1c5bf-421">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-421">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="1c5bf-422">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="1c5bf-422">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1c5bf-423">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="1c5bf-423">Platform</span></span></th>
    <th><span data-ttu-id="1c5bf-424">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="1c5bf-424">Extension points</span></span></th>
    <th><span data-ttu-id="1c5bf-425">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="1c5bf-425">API requirement sets</span></span></th>
    <th><span data-ttu-id="1c5bf-426"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-426"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-427">Office Online</span><span class="sxs-lookup"><span data-stu-id="1c5bf-427">Office Online</span></span></td>
    <td> <span data-ttu-id="1c5bf-428">- Contenu</span><span class="sxs-lookup"><span data-stu-id="1c5bf-428">- Content</span></span><br><span data-ttu-id="1c5bf-429">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-429">
         - Taskpane</span></span><br><span data-ttu-id="1c5bf-430">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-430">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c5bf-431">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-431">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c5bf-432">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1c5bf-432">-ActiveView</span></span><br><span data-ttu-id="1c5bf-433">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-433">
         -CompressedFile</span></span><br><span data-ttu-id="1c5bf-434">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-434">
         -DocumentEvents</span></span><br><span data-ttu-id="1c5bf-435">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c5bf-435">
         - File</span></span><br><span data-ttu-id="1c5bf-436">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-436">
         -ImageCoercion</span></span><br><span data-ttu-id="1c5bf-437">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-437">
         -PdfFile</span></span><br><span data-ttu-id="1c5bf-438">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c5bf-438">
         - Selection</span></span><br><span data-ttu-id="1c5bf-439">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-439">
         - Settings</span></span><br><span data-ttu-id="1c5bf-440">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-440">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-441">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1c5bf-441">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="1c5bf-442">- Contenu</span><span class="sxs-lookup"><span data-stu-id="1c5bf-442">- Content</span></span><br><span data-ttu-id="1c5bf-443">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-443">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="1c5bf-444">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="1c5bf-444">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="1c5bf-445">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1c5bf-445">-ActiveView</span></span><br><span data-ttu-id="1c5bf-446">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-446">
         -CompressedFile</span></span><br><span data-ttu-id="1c5bf-447">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-447">
         -DocumentEvents</span></span><br><span data-ttu-id="1c5bf-448">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c5bf-448">
         - File</span></span><br><span data-ttu-id="1c5bf-449">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-449">
         -ImageCoercion</span></span><br><span data-ttu-id="1c5bf-450">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-450">
         -PdfFile</span></span><br><span data-ttu-id="1c5bf-451">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c5bf-451">
         - Selection</span></span><br><span data-ttu-id="1c5bf-452">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-452">
         - Settings</span></span><br><span data-ttu-id="1c5bf-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-453">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-454">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="1c5bf-454">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="1c5bf-455">- Contenu</span><span class="sxs-lookup"><span data-stu-id="1c5bf-455">- Content</span></span><br><span data-ttu-id="1c5bf-456">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-456">
         - Taskpane</span></span><br><span data-ttu-id="1c5bf-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c5bf-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c5bf-459">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1c5bf-459">-ActiveView</span></span><br><span data-ttu-id="1c5bf-460">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-460">
         -CompressedFile</span></span><br><span data-ttu-id="1c5bf-461">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-461">
         -DocumentEvents</span></span><br><span data-ttu-id="1c5bf-462">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c5bf-462">
         - File</span></span><br><span data-ttu-id="1c5bf-463">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-463">
         -ImageCoercion</span></span><br><span data-ttu-id="1c5bf-464">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-464">
         -PdfFile</span></span><br><span data-ttu-id="1c5bf-465">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c5bf-465">
         - Selection</span></span><br><span data-ttu-id="1c5bf-466">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-466">
         - Settings</span></span><br><span data-ttu-id="1c5bf-467">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-467">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-468">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="1c5bf-468">Office for iOS</span></span></td>
    <td> <span data-ttu-id="1c5bf-469">- Contenu</span><span class="sxs-lookup"><span data-stu-id="1c5bf-469">- Content</span></span><br><span data-ttu-id="1c5bf-470">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-470">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="1c5bf-471">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-471">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="1c5bf-472">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1c5bf-472">-ActiveView</span></span><br><span data-ttu-id="1c5bf-473">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-473">
         -CompressedFile</span></span><br><span data-ttu-id="1c5bf-474">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-474">
         -DocumentEvents</span></span><br><span data-ttu-id="1c5bf-475">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c5bf-475">
         - File</span></span><br><span data-ttu-id="1c5bf-476">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-476">
         -PdfFile</span></span><br><span data-ttu-id="1c5bf-477">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c5bf-477">
         - Selection</span></span><br><span data-ttu-id="1c5bf-478">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-478">
         - Settings</span></span><br><span data-ttu-id="1c5bf-479">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-479">
         -TextCoercion</span></span><br><span data-ttu-id="1c5bf-480">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-480">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-481">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="1c5bf-481">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="1c5bf-482">- Contenu</span><span class="sxs-lookup"><span data-stu-id="1c5bf-482">- Content</span></span><br><span data-ttu-id="1c5bf-483">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-483">
         - Taskpane</span></span><br><span data-ttu-id="1c5bf-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c5bf-485">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-485">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c5bf-486">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1c5bf-486">-ActiveView</span></span><br><span data-ttu-id="1c5bf-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-487">
         -CompressedFile</span></span><br><span data-ttu-id="1c5bf-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-488">
         -DocumentEvents</span></span><br><span data-ttu-id="1c5bf-489">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c5bf-489">
         - File</span></span><br><span data-ttu-id="1c5bf-490">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-490">
         -ImageCoercion</span></span><br><span data-ttu-id="1c5bf-491">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c5bf-491">
         -PdfFile</span></span><br><span data-ttu-id="1c5bf-492">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c5bf-492">
         - Selection</span></span><br><span data-ttu-id="1c5bf-493">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-493">
         - Settings</span></span><br><span data-ttu-id="1c5bf-494">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-494">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="1c5bf-495">OneNote</span><span class="sxs-lookup"><span data-stu-id="1c5bf-495">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1c5bf-496">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="1c5bf-496">Platform</span></span></th>
    <th><span data-ttu-id="1c5bf-497">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="1c5bf-497">Extension points</span></span></th>
    <th><span data-ttu-id="1c5bf-498">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="1c5bf-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="1c5bf-499"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-499"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="1c5bf-500">Office Online</span><span class="sxs-lookup"><span data-stu-id="1c5bf-500">Office Online</span></span></td>
    <td> <span data-ttu-id="1c5bf-501">- Contenu</span><span class="sxs-lookup"><span data-stu-id="1c5bf-501">- Content</span></span><br><span data-ttu-id="1c5bf-502">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-502">
         - Taskpane</span></span><br><span data-ttu-id="1c5bf-503">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-503">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c5bf-504">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-504">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="1c5bf-505">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c5bf-505">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c5bf-506">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c5bf-506">-DocumentEvents</span></span><br><span data-ttu-id="1c5bf-507">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-507">
         -HtmlCoercion</span></span><br><span data-ttu-id="1c5bf-508">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-508">
         -ImageCoercion</span></span><br><span data-ttu-id="1c5bf-509">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c5bf-509">
         - Settings</span></span><br><span data-ttu-id="1c5bf-510">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c5bf-510">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="1c5bf-511">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="1c5bf-511">See also</span></span>

- [<span data-ttu-id="1c5bf-512">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-512">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="1c5bf-513">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="1c5bf-513">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="1c5bf-514">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="1c5bf-514">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="1c5bf-515">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-515">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)
