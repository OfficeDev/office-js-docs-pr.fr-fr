---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint et OneNote.
ms.date: 03/23/2018
ms.openlocfilehash: f50ab7e5312702eb25fbb2c8a25291c5ff5027a7
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19438871"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="76d88-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="76d88-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="76d88-104">Pour fonctionner comme prévu, il se peut que votre complément Office dépende d’un hôte Office spécifique, d’un ensemble de conditions requises, d’un membre d’API ou d’une version de l’API.</span><span class="sxs-lookup"><span data-stu-id="76d88-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="76d88-105">Les tableaux suivants contiennent la plateforme disponible, les points d’extension, les ensembles de conditions requises de l’API et les ensembles de conditions requises des API communes qui sont actuellement pris en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="76d88-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span> 

<span data-ttu-id="76d88-106">Si une cellule de tableau contient un astérisque (\*), cela signifie que nous travaillons sur celle-ci.</span><span class="sxs-lookup"><span data-stu-id="76d88-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="76d88-107">Pour les ensembles de conditions requises pour Project ou Access, consultez les [ensembles de conditions requises communs à Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="76d88-107">For requirement sets for Project or Access, see [Office common requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="76d88-p103">Le numéro de build pour Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que les ensembles de conditions requises des API communes, ExcelApi 1.1 et WordApi 1.1.</span><span class="sxs-lookup"><span data-stu-id="76d88-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="76d88-110">Excel</span><span class="sxs-lookup"><span data-stu-id="76d88-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="76d88-111">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="76d88-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="76d88-112">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="76d88-112">Extension points</span></span></th> 
    <th style="width:20%"><span data-ttu-id="76d88-113">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="76d88-113">API requirement sets</span></span></th> 
    <th style="width:40%"><span data-ttu-id="76d88-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="76d88-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="76d88-115">Office Online</span></span></td>
    <td> <span data-ttu-id="76d88-116">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="76d88-116">- Taskpane</span></span><br><span data-ttu-id="76d88-117">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="76d88-117">
        - Content</span></span><br><span data-ttu-id="76d88-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="76d88-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="76d88-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="76d88-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="76d88-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="76d88-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="76d88-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="76d88-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="76d88-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="76d88-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="76d88-124">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-124">
        -BindingEvents</span></span><br><span data-ttu-id="76d88-125">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-125">
        -DocumentEvents</span></span><br><span data-ttu-id="76d88-126">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-126">
        -MatrixBindings</span></span><br><span data-ttu-id="76d88-127">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-127">
        -MatrixCoercion</span></span><br><span data-ttu-id="76d88-128">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-128">
        -TableBindings</span></span><br><span data-ttu-id="76d88-129">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-129">
        -TableCoercion</span></span><br><span data-ttu-id="76d88-130">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-130">
        -TextBindings</span></span><br><span data-ttu-id="76d88-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="76d88-131">
        -CompressedFile</span></span><br><span data-ttu-id="76d88-132">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="76d88-132">
        - Settings</span></span><br><span data-ttu-id="76d88-133">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-133">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-134">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="76d88-134">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="76d88-135">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="76d88-135">
        - Taskpane</span></span><br><span data-ttu-id="76d88-136">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="76d88-136">
        - Content</span></span></td>
    <td>  <span data-ttu-id="76d88-137">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-137">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="76d88-138">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-138">
        -BindingEvents</span></span><br><span data-ttu-id="76d88-139">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-139">
        -DocumentEvents</span></span><br><span data-ttu-id="76d88-140">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-140">
        -MatrixBindings</span></span><br><span data-ttu-id="76d88-141">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-141">
        -MatrixCoercion</span></span><br><span data-ttu-id="76d88-142">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-142">
        -TableBindings</span></span><br><span data-ttu-id="76d88-143">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-143">
        -TableCoercion</span></span><br><span data-ttu-id="76d88-144">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-144">
        -TextBindings</span></span><br><span data-ttu-id="76d88-145">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="76d88-145">
        - Settings</span></span><br><span data-ttu-id="76d88-146">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-146">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-147">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="76d88-147">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="76d88-148">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="76d88-148">- Taskpane</span></span><br><span data-ttu-id="76d88-149">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="76d88-149">
        - Content</span></span><br><span data-ttu-id="76d88-150">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="76d88-150">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="76d88-151">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-151">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="76d88-152">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="76d88-152">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="76d88-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="76d88-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="76d88-154">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="76d88-154">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="76d88-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="76d88-156">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-156">-BindingEvents</span></span><br><span data-ttu-id="76d88-157">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-157">
        -DocumentEvents</span></span><br><span data-ttu-id="76d88-158">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-158">
        -MatrixBindings</span></span><br><span data-ttu-id="76d88-159">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-159">
        -MatrixCoercion</span></span><br><span data-ttu-id="76d88-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-160">
        -TableBindings</span></span><br><span data-ttu-id="76d88-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-161">
        -TableCoercion</span></span><br><span data-ttu-id="76d88-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-162">
        -TextBindings</span></span><br><span data-ttu-id="76d88-163">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="76d88-163">
        - Settings</span></span><br><span data-ttu-id="76d88-164">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-164">
        -TextCoercion</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-165">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="76d88-165">Office for iOS</span></span></td>
    <td><span data-ttu-id="76d88-166">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="76d88-166">- Taskpane</span></span><br><span data-ttu-id="76d88-167">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="76d88-167">
        - Content</span></span></td>
    <td><span data-ttu-id="76d88-168">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-168">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="76d88-169">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="76d88-169">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="76d88-170">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="76d88-170">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="76d88-171">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-171">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="76d88-172">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-172">-BindingEvents</span></span><br><span data-ttu-id="76d88-173">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-173">
        -DocumentEvents</span></span><br><span data-ttu-id="76d88-174">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-174">
        -MatrixBindings</span></span><br><span data-ttu-id="76d88-175">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-175">
        -MatrixCoercion</span></span><br><span data-ttu-id="76d88-176">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-176">
        -TableBindings</span></span><br><span data-ttu-id="76d88-177">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-177">
        -TableCoercion</span></span><br><span data-ttu-id="76d88-178">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-178">
        -TextBindings</span></span><br><span data-ttu-id="76d88-179">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="76d88-179">
        - Settings</span></span><br><span data-ttu-id="76d88-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-181">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="76d88-181">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="76d88-182">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="76d88-182">- Taskpane</span></span><br><span data-ttu-id="76d88-183">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="76d88-183">
        - Content</span></span><br><span data-ttu-id="76d88-184">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="76d88-184">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="76d88-185">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-185">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="76d88-186">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="76d88-186">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="76d88-187">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="76d88-187">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="76d88-188">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-188">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="76d88-189">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-189">-BindingEvents</span></span><br><span data-ttu-id="76d88-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-190">
        -DocumentEvents</span></span><br><span data-ttu-id="76d88-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-191">
        -MatrixBindings</span></span><br><span data-ttu-id="76d88-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-192">
        -MatrixCoercion</span></span><br><span data-ttu-id="76d88-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-193">
        -TableBindings</span></span><br><span data-ttu-id="76d88-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-194">
        -TableCoercion</span></span><br><span data-ttu-id="76d88-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-195">
        -TextBindings</span></span><br><span data-ttu-id="76d88-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-196">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="76d88-197">Outlook</span><span class="sxs-lookup"><span data-stu-id="76d88-197">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="76d88-198">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="76d88-198">Platform</span></span></th>
    <th><span data-ttu-id="76d88-199">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="76d88-199">Extension points</span></span></th> 
    <th><span data-ttu-id="76d88-200">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="76d88-200">API requirement sets</span></span></th> 
    <th><span data-ttu-id="76d88-201"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="76d88-201"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-202">Office Online</span><span class="sxs-lookup"><span data-stu-id="76d88-202">Office Online</span></span></td>
    <td> <span data-ttu-id="76d88-203">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="76d88-203">- Mail Read</span></span><br><span data-ttu-id="76d88-204">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="76d88-204">
      - Mail Compose</span></span><br><span data-ttu-id="76d88-205">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="76d88-205">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="76d88-206">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-206">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="76d88-207">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="76d88-207">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="76d88-208">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="76d88-208">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="76d88-209">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="76d88-209">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="76d88-210">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="76d88-210">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="76d88-211">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="76d88-211">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="76d88-212">non disponible</span><span class="sxs-lookup"><span data-stu-id="76d88-212">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-213">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="76d88-213">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="76d88-214">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="76d88-214">- Mail Read</span></span><br><span data-ttu-id="76d88-215">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="76d88-215">
      - Mail Compose</span></span><br><span data-ttu-id="76d88-216">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="76d88-216">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="76d88-217">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-217">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="76d88-218">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="76d88-218">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="76d88-219">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="76d88-219">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="76d88-220">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="76d88-220">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="76d88-221">non disponible</span><span class="sxs-lookup"><span data-stu-id="76d88-221">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-222">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="76d88-222">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="76d88-223">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="76d88-223">- Mail Read</span></span><br><span data-ttu-id="76d88-224">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="76d88-224">
      - Mail Compose</span></span><br><span data-ttu-id="76d88-225">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="76d88-225">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="76d88-226">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="76d88-226">
      - Modules</span></span></td>
    <td> <span data-ttu-id="76d88-227">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-227">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="76d88-228">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="76d88-228">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="76d88-229">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="76d88-229">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="76d88-230">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="76d88-230">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="76d88-231">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="76d88-231">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="76d88-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="76d88-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="76d88-233">non disponible</span><span class="sxs-lookup"><span data-stu-id="76d88-233">not available</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-234">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="76d88-234">Office for iOS</span></span></td>
    <td> <span data-ttu-id="76d88-235">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="76d88-235">- Mail Read</span></span><br><span data-ttu-id="76d88-236">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="76d88-236">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="76d88-237">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-237">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="76d88-238">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="76d88-238">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="76d88-239">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="76d88-239">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="76d88-240">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="76d88-240">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="76d88-241">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="76d88-241">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>    
    <td><span data-ttu-id="76d88-242">non disponible</span><span class="sxs-lookup"><span data-stu-id="76d88-242">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-243">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="76d88-243">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="76d88-244">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="76d88-244">- Mail Read</span></span><br><span data-ttu-id="76d88-245">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="76d88-245">
      - Mail Compose</span></span><br><span data-ttu-id="76d88-246">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="76d88-246">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="76d88-247">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-247">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="76d88-248">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="76d88-248">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="76d88-249">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="76d88-249">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="76d88-250">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="76d88-250">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="76d88-251">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="76d88-251">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="76d88-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="76d88-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="76d88-253">non disponible</span><span class="sxs-lookup"><span data-stu-id="76d88-253">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-254">Office pour Android</span><span class="sxs-lookup"><span data-stu-id="76d88-254">Office for Android</span></span></td>
    <td> <span data-ttu-id="76d88-255">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="76d88-255">- Mail Read</span></span><br><span data-ttu-id="76d88-256">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="76d88-256">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="76d88-257">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-257">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="76d88-258">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="76d88-258">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="76d88-259">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="76d88-259">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="76d88-260">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="76d88-260">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="76d88-261">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="76d88-261">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="76d88-262">non disponible</span><span class="sxs-lookup"><span data-stu-id="76d88-262">not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="76d88-263">Word</span><span class="sxs-lookup"><span data-stu-id="76d88-263">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="76d88-264">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="76d88-264">Platform</span></span></th>
    <th><span data-ttu-id="76d88-265">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="76d88-265">Extension points</span></span></th> 
    <th><span data-ttu-id="76d88-266">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="76d88-266">API requirement sets</span></span></th> 
    <th><span data-ttu-id="76d88-267"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="76d88-267"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-268">Office Online</span><span class="sxs-lookup"><span data-stu-id="76d88-268">Office Online</span></span></td>
    <td> <span data-ttu-id="76d88-269">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="76d88-269">- Taskpane</span></span><br><span data-ttu-id="76d88-270">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="76d88-270">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="76d88-271">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-271">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="76d88-272">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="76d88-272">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="76d88-273">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="76d88-273">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="76d88-274">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-274">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="76d88-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-275">-BindingEvents</span></span><br><span data-ttu-id="76d88-276">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="76d88-276">
         -</span></span><br><span data-ttu-id="76d88-277">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-277">
         -MatrixBindings</span></span><br><span data-ttu-id="76d88-278">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-278">
         -MatrixCoercion</span></span><br><span data-ttu-id="76d88-279">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-279">
         -TableBindings</span></span><br><span data-ttu-id="76d88-280">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-280">
         -TableCoercion</span></span><br><span data-ttu-id="76d88-281">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-281">
         -TextBindings</span></span><br><span data-ttu-id="76d88-282">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-282">
         -DocumentEvents</span></span><br><span data-ttu-id="76d88-283">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="76d88-283">
         -TextFile</span></span><br><span data-ttu-id="76d88-284">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-284">
         -ImageCoercion</span></span><br><span data-ttu-id="76d88-285">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="76d88-285">
         - Settings</span></span><br><span data-ttu-id="76d88-286">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-286">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-287">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="76d88-287">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="76d88-288">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="76d88-288">- Taskpane</span></span></td>
    <td> <span data-ttu-id="76d88-289">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-289">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="76d88-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-290">-BindingEvents</span></span><br><span data-ttu-id="76d88-291">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="76d88-291">
         -CompressedFile</span></span><br><span data-ttu-id="76d88-292">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="76d88-292">
         -CustomXmlPart</span></span><br><span data-ttu-id="76d88-293">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-293">
         -DocumentEvents</span></span><br><span data-ttu-id="76d88-294">
         - File</span><span class="sxs-lookup"><span data-stu-id="76d88-294">
         - File</span></span><br><span data-ttu-id="76d88-295">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-295">
         -HtmlCoercion</span></span><br><span data-ttu-id="76d88-296">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-296">
         -ImageCoercion</span></span><br><span data-ttu-id="76d88-297">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-297">
         -OoxmlCoercion</span></span><br><span data-ttu-id="76d88-298">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-298">
         -TableBindings</span></span><br><span data-ttu-id="76d88-299">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-299">
         -TableCoercion</span></span><br><span data-ttu-id="76d88-300">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-300">
         -TextBindings</span></span><br><span data-ttu-id="76d88-301">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="76d88-301">
         -TextFile</span></span><br><span data-ttu-id="76d88-302">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="76d88-302">
         - Settings</span></span><br><span data-ttu-id="76d88-303">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-303">
         -TextCoercion</span></span><br><span data-ttu-id="76d88-304">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-304">
         -MatrixCoercion</span></span><br><span data-ttu-id="76d88-305">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-305">
         - Matrix Bindings</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-306">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="76d88-306">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="76d88-307">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="76d88-307">- Taskpane</span></span><br><span data-ttu-id="76d88-308">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="76d88-308">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="76d88-309">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-309">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="76d88-310">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="76d88-310">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="76d88-311">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="76d88-311">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="76d88-312">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-312">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="76d88-313">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-313">-BindingEvents</span></span><br><span data-ttu-id="76d88-314">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="76d88-314">
         -CompressedFile</span></span><br><span data-ttu-id="76d88-315">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="76d88-315">
         -CustomXmlPart</span></span><br><span data-ttu-id="76d88-316">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-316">
         -DocumentEvents</span></span><br><span data-ttu-id="76d88-317">
         - File</span><span class="sxs-lookup"><span data-stu-id="76d88-317">
         - File</span></span><br><span data-ttu-id="76d88-318">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-318">
         -HtmlCoercion</span></span><br><span data-ttu-id="76d88-319">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-319">
         -ImageCoercion</span></span><br><span data-ttu-id="76d88-320">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-320">
         -OoxmlCoercion</span></span><br><span data-ttu-id="76d88-321">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-321">
         -TableBindings</span></span><br><span data-ttu-id="76d88-322">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-322">
         -TableCoercion</span></span><br><span data-ttu-id="76d88-323">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-323">
         -TextBindings</span></span><br><span data-ttu-id="76d88-324">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="76d88-324">
         -TextFile</span></span><br><span data-ttu-id="76d88-325">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="76d88-325">
         - Settings</span></span><br><span data-ttu-id="76d88-326">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-326">
         -TextCoercion</span></span><br><span data-ttu-id="76d88-327">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-327">
         -MatrixCoercion</span></span><br><span data-ttu-id="76d88-328">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-328">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-329">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="76d88-329">Office for iOS</span></span></td>
    <td> <span data-ttu-id="76d88-330">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="76d88-330">- Taskpane</span></span></td>
    <td> <span data-ttu-id="76d88-331">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-331">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="76d88-332">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="76d88-332">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="76d88-333">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="76d88-333">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="76d88-334">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="76d88-334">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="76d88-335">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-335">-BindingEvents</span></span><br><span data-ttu-id="76d88-336">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="76d88-336">
         -CompressedFile</span></span><br><span data-ttu-id="76d88-337">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="76d88-337">
         -CustomXmlPart</span></span><br><span data-ttu-id="76d88-338">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-338">
         -DocumentEvents</span></span><br><span data-ttu-id="76d88-339">
         - File</span><span class="sxs-lookup"><span data-stu-id="76d88-339">
         - File</span></span><br><span data-ttu-id="76d88-340">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-340">
         -HtmlCoercion</span></span><br><span data-ttu-id="76d88-341">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-341">
         -ImageCoercion</span></span><br><span data-ttu-id="76d88-342">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-342">
         -OoxmlCoercion</span></span><br><span data-ttu-id="76d88-343">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-343">
         -TableBindings</span></span><br><span data-ttu-id="76d88-344">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-344">
         -TableCoercion</span></span><br><span data-ttu-id="76d88-345">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-345">
         -TextBindings</span></span><br><span data-ttu-id="76d88-346">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="76d88-346">
         -TextFile</span></span><br><span data-ttu-id="76d88-347">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="76d88-347">
         - Settings</span></span><br><span data-ttu-id="76d88-348">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-348">
         -TextCoercion</span></span><br><span data-ttu-id="76d88-349">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-349">
         -MatrixCoercion</span></span><br><span data-ttu-id="76d88-350">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-350">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-351">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="76d88-351">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="76d88-352">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="76d88-352">- Taskpane</span></span><br><span data-ttu-id="76d88-353">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="76d88-353">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="76d88-354">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-354">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="76d88-355">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="76d88-355">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="76d88-356">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="76d88-356">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="76d88-357">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="76d88-357">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="76d88-358">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-358">-BindingEvents</span></span><br><span data-ttu-id="76d88-359">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="76d88-359">
         -CompressedFile</span></span><br><span data-ttu-id="76d88-360">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="76d88-360">
         -CustomXmlPart</span></span><br><span data-ttu-id="76d88-361">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-361">
         -DocumentEvents</span></span><br><span data-ttu-id="76d88-362">
         - File</span><span class="sxs-lookup"><span data-stu-id="76d88-362">
         - File</span></span><br><span data-ttu-id="76d88-363">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-363">
         -HtmlCoercion</span></span><br><span data-ttu-id="76d88-364">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-364">
         -ImageCoercion</span></span><br><span data-ttu-id="76d88-365">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-365">
         -OoxmlCoercion</span></span><br><span data-ttu-id="76d88-366">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-366">
         -TableBindings</span></span><br><span data-ttu-id="76d88-367">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-367">
         -TableCoercion</span></span><br><span data-ttu-id="76d88-368">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-368">
         -TextBindings</span></span><br><span data-ttu-id="76d88-369">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="76d88-369">
         -TextFile</span></span><br><span data-ttu-id="76d88-370">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="76d88-370">
         - Settings</span></span><br><span data-ttu-id="76d88-371">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-371">
         -TextCoercion</span></span><br><span data-ttu-id="76d88-372">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-372">
         -MatrixCoercion</span></span><br><span data-ttu-id="76d88-373">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="76d88-373">
         - Matrix Bindings</span></span> </td> 
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="76d88-374">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="76d88-374">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="76d88-375">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="76d88-375">Platform</span></span></th>
    <th><span data-ttu-id="76d88-376">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="76d88-376">Extension points</span></span></th> 
    <th><span data-ttu-id="76d88-377">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="76d88-377">API requirement sets</span></span></th> 
    <th><span data-ttu-id="76d88-378"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="76d88-378"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-379">Office Online</span><span class="sxs-lookup"><span data-stu-id="76d88-379">Office Online</span></span></td>
    <td> <span data-ttu-id="76d88-380">- Contenu</span><span class="sxs-lookup"><span data-stu-id="76d88-380">- Content</span></span><br><span data-ttu-id="76d88-381">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="76d88-381">
         - Taskpane</span></span><br><span data-ttu-id="76d88-382">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="76d88-382">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="76d88-383">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-383">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="76d88-384">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="76d88-384">-ActiveView</span></span><br><span data-ttu-id="76d88-385">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="76d88-385">
         -CompressedFile</span></span><br><span data-ttu-id="76d88-386">
         - File</span><span class="sxs-lookup"><span data-stu-id="76d88-386">
         - File</span></span><br><span data-ttu-id="76d88-387">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="76d88-387">
         - Selection</span></span><br><span data-ttu-id="76d88-388">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="76d88-388">
         - Settings</span></span><br><span data-ttu-id="76d88-389">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-389">
         -TextCoercion</span></span><br><span data-ttu-id="76d88-390">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-390">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-391">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="76d88-391">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="76d88-392">- Contenu</span><span class="sxs-lookup"><span data-stu-id="76d88-392">- Content</span></span><br><span data-ttu-id="76d88-393">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="76d88-393">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="76d88-394">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="76d88-394">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="76d88-395">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="76d88-395">-ActiveView</span></span><br><span data-ttu-id="76d88-396">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="76d88-396">
         -CompressedFile</span></span><br><span data-ttu-id="76d88-397">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-397">
         -DocumentEvents</span></span><br><span data-ttu-id="76d88-398">
         - File</span><span class="sxs-lookup"><span data-stu-id="76d88-398">
         - File</span></span><br><span data-ttu-id="76d88-399">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="76d88-399">
         - Selection</span></span><br><span data-ttu-id="76d88-400">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="76d88-400">
         - Settings</span></span><br><span data-ttu-id="76d88-401">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-401">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-402">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="76d88-402">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="76d88-403">- Contenu</span><span class="sxs-lookup"><span data-stu-id="76d88-403">- Content</span></span><br><span data-ttu-id="76d88-404">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="76d88-404">
         - Taskpane</span></span><br><span data-ttu-id="76d88-405">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="76d88-405">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="76d88-406">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-406">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="76d88-407">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="76d88-407">-ActiveView</span></span><br><span data-ttu-id="76d88-408">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="76d88-408">
         -CompressedFile</span></span><br><span data-ttu-id="76d88-409">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-409">
         -DocumentEvents</span></span><br><span data-ttu-id="76d88-410">
         - File</span><span class="sxs-lookup"><span data-stu-id="76d88-410">
         - File</span></span><br><span data-ttu-id="76d88-411">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="76d88-411">
         - Selection</span></span><br><span data-ttu-id="76d88-412">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="76d88-412">
         - Settings</span></span><br><span data-ttu-id="76d88-413">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-413">
         -TextCoercion</span></span><br><span data-ttu-id="76d88-414">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-414">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-415">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="76d88-415">Office for iOS</span></span></td>
    <td> <span data-ttu-id="76d88-416">- Contenu</span><span class="sxs-lookup"><span data-stu-id="76d88-416">- Content</span></span><br><span data-ttu-id="76d88-417">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="76d88-417">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="76d88-418">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-418">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="76d88-419">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="76d88-419">-ActiveView</span></span><br><span data-ttu-id="76d88-420">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="76d88-420">
         -CompressedFile</span></span><br><span data-ttu-id="76d88-421">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-421">
         -DocumentEvents</span></span><br><span data-ttu-id="76d88-422">
         - File</span><span class="sxs-lookup"><span data-stu-id="76d88-422">
         - File</span></span><br><span data-ttu-id="76d88-423">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="76d88-423">
         - Selection</span></span><br><span data-ttu-id="76d88-424">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="76d88-424">
         - Settings</span></span><br><span data-ttu-id="76d88-425">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-425">
         -TextCoercion</span></span><br><span data-ttu-id="76d88-426">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-426">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-427">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="76d88-427">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="76d88-428">- Contenu</span><span class="sxs-lookup"><span data-stu-id="76d88-428">- Content</span></span><br><span data-ttu-id="76d88-429">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="76d88-429">
         - Taskpane</span></span><br><span data-ttu-id="76d88-430">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="76d88-430">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="76d88-431">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-431">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="76d88-432">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="76d88-432">-ActiveView</span></span><br><span data-ttu-id="76d88-433">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="76d88-433">
         -CompressedFile</span></span><br><span data-ttu-id="76d88-434">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-434">
         -DocumentEvents</span></span><br><span data-ttu-id="76d88-435">
         - File</span><span class="sxs-lookup"><span data-stu-id="76d88-435">
         - File</span></span><br><span data-ttu-id="76d88-436">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="76d88-436">
         - Selection</span></span><br><span data-ttu-id="76d88-437">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="76d88-437">
         - Settings</span></span><br><span data-ttu-id="76d88-438">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-438">
         -TextCoercion</span></span><br><span data-ttu-id="76d88-439">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-439">
         -ImageCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="76d88-440">OneNote</span><span class="sxs-lookup"><span data-stu-id="76d88-440">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="76d88-441">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="76d88-441">Platform</span></span></th>
    <th><span data-ttu-id="76d88-442">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="76d88-442">Extension points</span></span></th> 
    <th><span data-ttu-id="76d88-443">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="76d88-443">API requirement sets</span></span></th> 
    <th><span data-ttu-id="76d88-444"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="76d88-444"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-445">Office Online</span><span class="sxs-lookup"><span data-stu-id="76d88-445">Office Online</span></span></td>
    <td> <span data-ttu-id="76d88-446">- Contenu</span><span class="sxs-lookup"><span data-stu-id="76d88-446">- Content</span></span><br><span data-ttu-id="76d88-447">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="76d88-447">
         - Taskpane</span></span><br><span data-ttu-id="76d88-448">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="76d88-448">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="76d88-449">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-449">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="76d88-450">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="76d88-450">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="76d88-451">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="76d88-451">-DocumentEvents</span></span><br><span data-ttu-id="76d88-452">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="76d88-452">
         - Settings</span></span><br><span data-ttu-id="76d88-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-453">
         -TextCoercion</span></span><br><span data-ttu-id="76d88-454">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-454">
         -HtmlCoercion</span></span><br><span data-ttu-id="76d88-455">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="76d88-455">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-456">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="76d88-456">Office 2013 for Windows</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr> 
  <tr>
    <td><span data-ttu-id="76d88-457">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="76d88-457">Office 2016 for Windows</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-458">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="76d88-458">Office for iOS</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
  <tr>
    <td><span data-ttu-id="76d88-459">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="76d88-459">Office 2016 for Mac</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
</table>

<br/>

<span data-ttu-id="76d88-460">\* = Nous y travaillons.</span><span class="sxs-lookup"><span data-stu-id="76d88-460">\* = We're working on it.</span></span> 

## <a name="see-also"></a><span data-ttu-id="76d88-461">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="76d88-461">See also</span></span>

- [<span data-ttu-id="76d88-462">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="76d88-462">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="76d88-463">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="76d88-463">Common API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="76d88-464">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="76d88-464">Add-in Commands requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="76d88-465">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="76d88-465">JavaScript API for Office reference</span></span>](https://dev.office.com/reference/add-ins/javascript-api-for-office)

