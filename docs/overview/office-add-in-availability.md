---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint et OneNote.
ms.date: 07/31/2018
ms.openlocfilehash: 084029c0a5b70b73eaa0b3fcc180f4a813fb8b72
ms.sourcegitcommit: bc68b4cf811b45e8b8d1cbd7c8d2867359ab671b
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/02/2018
ms.locfileid: "21703909"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="c86cb-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="c86cb-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="c86cb-104">Pour fonctionner comme prévu, il se peut que votre complément Office dépende d’un hôte Office spécifique, d’un ensemble de conditions requises, d’un membre d’API ou d’une version de l’API.</span><span class="sxs-lookup"><span data-stu-id="c86cb-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="c86cb-105">Les tableaux suivants contiennent la plateforme disponible, les points d’extension, les ensembles de conditions requises de l’API et les ensembles de conditions requises des API communes qui sont actuellement pris en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="c86cb-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span> 

<span data-ttu-id="c86cb-106">Si une cellule de tableau contient un astérisque (\*), cela signifie que nous travaillons sur celle-ci.</span><span class="sxs-lookup"><span data-stu-id="c86cb-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="c86cb-107">Pour les ensembles de conditions requises pour Project ou Access, consultez les [ensembles de conditions requises communs à Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="c86cb-107">For requirement sets for Project or Access, see [Office common requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="c86cb-p103">Le numéro de build pour Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que les ensembles de conditions requises des API communes, ExcelApi 1.1 et WordApi 1.1.</span><span class="sxs-lookup"><span data-stu-id="c86cb-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="c86cb-110">Excel</span><span class="sxs-lookup"><span data-stu-id="c86cb-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="c86cb-111">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="c86cb-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="c86cb-112">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="c86cb-112">Extension points</span></span></th> 
    <th style="width:20%"><span data-ttu-id="c86cb-113">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="c86cb-113">API requirement sets</span></span></th> 
    <th style="width:40%"><span data-ttu-id="c86cb-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="c86cb-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="c86cb-115">Office Online</span></span></td>
    <td> <span data-ttu-id="c86cb-116">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c86cb-116">- Taskpane</span></span><br><span data-ttu-id="c86cb-117">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="c86cb-117">
        - Content</span></span><br><span data-ttu-id="c86cb-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="c86cb-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="c86cb-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c86cb-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c86cb-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c86cb-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c86cb-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c86cb-124">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-124">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c86cb-125">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="c86cb-126">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-126">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c86cb-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-127">
        -BindingEvents</span></span><br><span data-ttu-id="c86cb-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-128">
        -DocumentEvents</span></span><br><span data-ttu-id="c86cb-129">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-129">
        -MatrixBindings</span></span><br><span data-ttu-id="c86cb-130">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-130">
        -MatrixCoercion</span></span><br><span data-ttu-id="c86cb-131">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-131">
        -TableBindings</span></span><br><span data-ttu-id="c86cb-132">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-132">
        -TableCoercion</span></span><br><span data-ttu-id="c86cb-133">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-133">
        -TextBindings</span></span><br><span data-ttu-id="c86cb-134">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c86cb-134">
        -CompressedFile</span></span><br><span data-ttu-id="c86cb-135">
        - Paramètres</span><span class="sxs-lookup"><span data-stu-id="c86cb-135">
        - Settings</span></span><br><span data-ttu-id="c86cb-136">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-136">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-137">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="c86cb-137">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="c86cb-138">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="c86cb-138">
        - Taskpane</span></span><br><span data-ttu-id="c86cb-139">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="c86cb-139">
        - Content</span></span></td>
    <td>  <span data-ttu-id="c86cb-140">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-140">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c86cb-141">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-141">
        -BindingEvents</span></span><br><span data-ttu-id="c86cb-142">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-142">
        -DocumentEvents</span></span><br><span data-ttu-id="c86cb-143">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-143">
        -MatrixBindings</span></span><br><span data-ttu-id="c86cb-144">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-144">
        -MatrixCoercion</span></span><br><span data-ttu-id="c86cb-145">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-145">
        -TableBindings</span></span><br><span data-ttu-id="c86cb-146">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-146">
        -TableCoercion</span></span><br><span data-ttu-id="c86cb-147">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-147">
        -TextBindings</span></span><br><span data-ttu-id="c86cb-148">
        - Paramètres</span><span class="sxs-lookup"><span data-stu-id="c86cb-148">
        - Settings</span></span><br><span data-ttu-id="c86cb-149">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-149">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-150">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="c86cb-150">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="c86cb-151">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c86cb-151">- Taskpane</span></span><br><span data-ttu-id="c86cb-152">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="c86cb-152">
        - Content</span></span><br><span data-ttu-id="c86cb-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c86cb-154">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-154">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c86cb-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c86cb-156">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-156">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c86cb-157">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-157">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c86cb-158">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-158">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c86cb-159">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-159">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c86cb-160">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-160">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="c86cb-161">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-161">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c86cb-162">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-162">-BindingEvents</span></span><br><span data-ttu-id="c86cb-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-163">
        -DocumentEvents</span></span><br><span data-ttu-id="c86cb-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-164">
        -MatrixBindings</span></span><br><span data-ttu-id="c86cb-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-165">
        -MatrixCoercion</span></span><br><span data-ttu-id="c86cb-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-166">
        -TableBindings</span></span><br><span data-ttu-id="c86cb-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-167">
        -TableCoercion</span></span><br><span data-ttu-id="c86cb-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-168">
        -TextBindings</span></span><br><span data-ttu-id="c86cb-169">
        - Paramètres</span><span class="sxs-lookup"><span data-stu-id="c86cb-169">
        - Settings</span></span><br><span data-ttu-id="c86cb-170">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-170">
        -TextCoercion</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-171">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="c86cb-171">Office for iOS</span></span></td>
    <td><span data-ttu-id="c86cb-172">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c86cb-172">- Taskpane</span></span><br><span data-ttu-id="c86cb-173">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="c86cb-173">
        - Content</span></span></td>
    <td><span data-ttu-id="c86cb-174">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-174">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c86cb-175">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-175">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c86cb-176">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-176">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c86cb-177">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-177">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c86cb-178">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-178">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c86cb-179">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-179">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c86cb-180">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-180">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="c86cb-181">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-181">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c86cb-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-182">-BindingEvents</span></span><br><span data-ttu-id="c86cb-183">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-183">
        -DocumentEvents</span></span><br><span data-ttu-id="c86cb-184">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-184">
        -MatrixBindings</span></span><br><span data-ttu-id="c86cb-185">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-185">
        -MatrixCoercion</span></span><br><span data-ttu-id="c86cb-186">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-186">
        -TableBindings</span></span><br><span data-ttu-id="c86cb-187">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-187">
        -TableCoercion</span></span><br><span data-ttu-id="c86cb-188">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-188">
        -TextBindings</span></span><br><span data-ttu-id="c86cb-189">
        - Paramètres</span><span class="sxs-lookup"><span data-stu-id="c86cb-189">
        - Settings</span></span><br><span data-ttu-id="c86cb-190">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-190">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-191">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="c86cb-191">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="c86cb-192">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c86cb-192">- Taskpane</span></span><br><span data-ttu-id="c86cb-193">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="c86cb-193">
        - Content</span></span><br><span data-ttu-id="c86cb-194">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-194">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c86cb-195">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-195">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c86cb-196">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-196">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c86cb-197">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-197">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c86cb-198">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-198">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c86cb-199">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-199">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c86cb-200">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-200">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c86cb-201">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-201">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="c86cb-202">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-202">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c86cb-203">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-203">-BindingEvents</span></span><br><span data-ttu-id="c86cb-204">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-204">
        -DocumentEvents</span></span><br><span data-ttu-id="c86cb-205">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-205">
        -MatrixBindings</span></span><br><span data-ttu-id="c86cb-206">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-206">
        -MatrixCoercion</span></span><br><span data-ttu-id="c86cb-207">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-207">
        -TableBindings</span></span><br><span data-ttu-id="c86cb-208">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-208">
        -TableCoercion</span></span><br><span data-ttu-id="c86cb-209">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-209">
        -TextBindings</span></span><br><span data-ttu-id="c86cb-210">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-210">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="c86cb-211">Outlook</span><span class="sxs-lookup"><span data-stu-id="c86cb-211">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c86cb-212">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="c86cb-212">Platform</span></span></th>
    <th><span data-ttu-id="c86cb-213">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="c86cb-213">Extension points</span></span></th> 
    <th><span data-ttu-id="c86cb-214">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="c86cb-214">API requirement sets</span></span></th> 
    <th><span data-ttu-id="c86cb-215"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="c86cb-215"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-216">Office Online</span><span class="sxs-lookup"><span data-stu-id="c86cb-216">Office Online</span></span></td>
    <td> <span data-ttu-id="c86cb-217">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="c86cb-217">- Mail Read</span></span><br><span data-ttu-id="c86cb-218">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="c86cb-218">
      - Mail Compose</span></span><br><span data-ttu-id="c86cb-219">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-219">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c86cb-220">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-220">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c86cb-221">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-221">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c86cb-222">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-222">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c86cb-223">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-223">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c86cb-224">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-224">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c86cb-225">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-225">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c86cb-226">Non disponible</span><span class="sxs-lookup"><span data-stu-id="c86cb-226">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-227">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="c86cb-227">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="c86cb-228">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="c86cb-228">- Mail Read</span></span><br><span data-ttu-id="c86cb-229">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="c86cb-229">
      - Mail Compose</span></span><br><span data-ttu-id="c86cb-230">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-230">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c86cb-231">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-231">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c86cb-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c86cb-233">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-233">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c86cb-234">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-234">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="c86cb-235">Non disponible</span><span class="sxs-lookup"><span data-stu-id="c86cb-235">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-236">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="c86cb-236">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="c86cb-237">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="c86cb-237">- Mail Read</span></span><br><span data-ttu-id="c86cb-238">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="c86cb-238">
      - Mail Compose</span></span><br><span data-ttu-id="c86cb-239">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-239">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c86cb-240">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="c86cb-240">
      - Modules</span></span></td>
    <td> <span data-ttu-id="c86cb-241">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-241">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c86cb-242">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-242">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c86cb-243">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-243">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c86cb-244">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-244">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c86cb-245">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-245">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c86cb-246">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-246">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c86cb-247">Non disponible</span><span class="sxs-lookup"><span data-stu-id="c86cb-247">Not available</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-248">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="c86cb-248">Office for iOS</span></span></td>
    <td> <span data-ttu-id="c86cb-249">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="c86cb-249">- Mail Read</span></span><br><span data-ttu-id="c86cb-250">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-250">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c86cb-251">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-251">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c86cb-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c86cb-253">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-253">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c86cb-254">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-254">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c86cb-255">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-255">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>    
    <td><span data-ttu-id="c86cb-256">Non disponible</span><span class="sxs-lookup"><span data-stu-id="c86cb-256">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-257">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="c86cb-257">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="c86cb-258">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="c86cb-258">- Mail Read</span></span><br><span data-ttu-id="c86cb-259">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="c86cb-259">
      - Mail Compose</span></span><br><span data-ttu-id="c86cb-260">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-260">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c86cb-261">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-261">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c86cb-262">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-262">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c86cb-263">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-263">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c86cb-264">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-264">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c86cb-265">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-265">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c86cb-266">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-266">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c86cb-267">Non disponible</span><span class="sxs-lookup"><span data-stu-id="c86cb-267">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-268">Office pour Android</span><span class="sxs-lookup"><span data-stu-id="c86cb-268">Office for Android</span></span></td>
    <td> <span data-ttu-id="c86cb-269">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="c86cb-269">- Mail Read</span></span><br><span data-ttu-id="c86cb-270">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-270">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c86cb-271">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-271">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c86cb-272">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-272">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c86cb-273">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-273">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c86cb-274">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-274">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c86cb-275">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-275">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c86cb-276">Non disponible</span><span class="sxs-lookup"><span data-stu-id="c86cb-276">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="c86cb-277">Word</span><span class="sxs-lookup"><span data-stu-id="c86cb-277">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c86cb-278">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="c86cb-278">Platform</span></span></th>
    <th><span data-ttu-id="c86cb-279">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="c86cb-279">Extension points</span></span></th> 
    <th><span data-ttu-id="c86cb-280">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="c86cb-280">API requirement sets</span></span></th> 
    <th><span data-ttu-id="c86cb-281"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="c86cb-281"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-282">Office Online</span><span class="sxs-lookup"><span data-stu-id="c86cb-282">Office Online</span></span></td>
    <td> <span data-ttu-id="c86cb-283">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c86cb-283">- Taskpane</span></span><br><span data-ttu-id="c86cb-284">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-284">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c86cb-285">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-285">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c86cb-286">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-286">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c86cb-287">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-287">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c86cb-288">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-288">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c86cb-289">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-289">-BindingEvents</span></span><br><span data-ttu-id="c86cb-290">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c86cb-290">
         -</span></span><br><span data-ttu-id="c86cb-291">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-291">
         -MatrixBindings</span></span><br><span data-ttu-id="c86cb-292">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-292">
         -MatrixCoercion</span></span><br><span data-ttu-id="c86cb-293">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-293">
         -TableBindings</span></span><br><span data-ttu-id="c86cb-294">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-294">
         -TableCoercion</span></span><br><span data-ttu-id="c86cb-295">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-295">
         -TextBindings</span></span><br><span data-ttu-id="c86cb-296">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-296">
         -DocumentEvents</span></span><br><span data-ttu-id="c86cb-297">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c86cb-297">
         -TextFile</span></span><br><span data-ttu-id="c86cb-298">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-298">
         -ImageCoercion</span></span><br><span data-ttu-id="c86cb-299">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="c86cb-299">
         - Settings</span></span><br><span data-ttu-id="c86cb-300">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-300">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-301">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="c86cb-301">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="c86cb-302">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c86cb-302">- Taskpane</span></span></td>
    <td> <span data-ttu-id="c86cb-303">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-303">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c86cb-304">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-304">-BindingEvents</span></span><br><span data-ttu-id="c86cb-305">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c86cb-305">
         -CompressedFile</span></span><br><span data-ttu-id="c86cb-306">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="c86cb-306">
         -CustomXmlPart</span></span><br><span data-ttu-id="c86cb-307">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-307">
         -DocumentEvents</span></span><br><span data-ttu-id="c86cb-308">
         - File</span><span class="sxs-lookup"><span data-stu-id="c86cb-308">
         - File</span></span><br><span data-ttu-id="c86cb-309">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-309">
         -HtmlCoercion</span></span><br><span data-ttu-id="c86cb-310">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-310">
         -ImageCoercion</span></span><br><span data-ttu-id="c86cb-311">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-311">
         -OoxmlCoercion</span></span><br><span data-ttu-id="c86cb-312">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-312">
         -TableBindings</span></span><br><span data-ttu-id="c86cb-313">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-313">
         -TableCoercion</span></span><br><span data-ttu-id="c86cb-314">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-314">
         -TextBindings</span></span><br><span data-ttu-id="c86cb-315">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c86cb-315">
         -TextFile</span></span><br><span data-ttu-id="c86cb-316">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="c86cb-316">
         - Settings</span></span><br><span data-ttu-id="c86cb-317">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-317">
         -TextCoercion</span></span><br><span data-ttu-id="c86cb-318">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-318">
         -MatrixCoercion</span></span><br><span data-ttu-id="c86cb-319">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-319">
         - Matrix Bindings</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-320">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="c86cb-320">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="c86cb-321">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c86cb-321">- Taskpane</span></span><br><span data-ttu-id="c86cb-322">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-322">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c86cb-323">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-323">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c86cb-324">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-324">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c86cb-325">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-325">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c86cb-326">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-326">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c86cb-327">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-327">-BindingEvents</span></span><br><span data-ttu-id="c86cb-328">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c86cb-328">
         -CompressedFile</span></span><br><span data-ttu-id="c86cb-329">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="c86cb-329">
         -CustomXmlPart</span></span><br><span data-ttu-id="c86cb-330">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-330">
         -DocumentEvents</span></span><br><span data-ttu-id="c86cb-331">
         - File</span><span class="sxs-lookup"><span data-stu-id="c86cb-331">
         - File</span></span><br><span data-ttu-id="c86cb-332">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-332">
         -HtmlCoercion</span></span><br><span data-ttu-id="c86cb-333">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-333">
         -ImageCoercion</span></span><br><span data-ttu-id="c86cb-334">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-334">
         -OoxmlCoercion</span></span><br><span data-ttu-id="c86cb-335">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-335">
         -TableBindings</span></span><br><span data-ttu-id="c86cb-336">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-336">
         -TableCoercion</span></span><br><span data-ttu-id="c86cb-337">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-337">
         -TextBindings</span></span><br><span data-ttu-id="c86cb-338">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c86cb-338">
         -TextFile</span></span><br><span data-ttu-id="c86cb-339">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="c86cb-339">
         - Settings</span></span><br><span data-ttu-id="c86cb-340">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-340">
         -TextCoercion</span></span><br><span data-ttu-id="c86cb-341">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-341">
         -MatrixCoercion</span></span><br><span data-ttu-id="c86cb-342">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-342">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-343">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="c86cb-343">Office for iOS</span></span></td>
    <td> <span data-ttu-id="c86cb-344">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c86cb-344">- Taskpane</span></span></td>
    <td> <span data-ttu-id="c86cb-345">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-345">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c86cb-346">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-346">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c86cb-347">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-347">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c86cb-348">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="c86cb-348">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="c86cb-349">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-349">-BindingEvents</span></span><br><span data-ttu-id="c86cb-350">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c86cb-350">
         -CompressedFile</span></span><br><span data-ttu-id="c86cb-351">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="c86cb-351">
         -CustomXmlPart</span></span><br><span data-ttu-id="c86cb-352">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-352">
         -DocumentEvents</span></span><br><span data-ttu-id="c86cb-353">
         - File</span><span class="sxs-lookup"><span data-stu-id="c86cb-353">
         - File</span></span><br><span data-ttu-id="c86cb-354">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-354">
         -HtmlCoercion</span></span><br><span data-ttu-id="c86cb-355">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-355">
         -ImageCoercion</span></span><br><span data-ttu-id="c86cb-356">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-356">
         -OoxmlCoercion</span></span><br><span data-ttu-id="c86cb-357">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-357">
         -TableBindings</span></span><br><span data-ttu-id="c86cb-358">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-358">
         -TableCoercion</span></span><br><span data-ttu-id="c86cb-359">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-359">
         -TextBindings</span></span><br><span data-ttu-id="c86cb-360">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c86cb-360">
         -TextFile</span></span><br><span data-ttu-id="c86cb-361">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="c86cb-361">
         - Settings</span></span><br><span data-ttu-id="c86cb-362">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-362">
         -TextCoercion</span></span><br><span data-ttu-id="c86cb-363">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-363">
         -MatrixCoercion</span></span><br><span data-ttu-id="c86cb-364">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-364">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-365">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="c86cb-365">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="c86cb-366">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="c86cb-366">- Taskpane</span></span><br><span data-ttu-id="c86cb-367">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-367">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c86cb-368">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-368">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c86cb-369">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-369">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c86cb-370">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-370">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c86cb-371">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="c86cb-371">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="c86cb-372">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-372">-BindingEvents</span></span><br><span data-ttu-id="c86cb-373">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c86cb-373">
         -CompressedFile</span></span><br><span data-ttu-id="c86cb-374">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="c86cb-374">
         -CustomXmlPart</span></span><br><span data-ttu-id="c86cb-375">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-375">
         -DocumentEvents</span></span><br><span data-ttu-id="c86cb-376">
         - File</span><span class="sxs-lookup"><span data-stu-id="c86cb-376">
         - File</span></span><br><span data-ttu-id="c86cb-377">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-377">
         -HtmlCoercion</span></span><br><span data-ttu-id="c86cb-378">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-378">
         -ImageCoercion</span></span><br><span data-ttu-id="c86cb-379">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-379">
         -OoxmlCoercion</span></span><br><span data-ttu-id="c86cb-380">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-380">
         -TableBindings</span></span><br><span data-ttu-id="c86cb-381">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-381">
         -TableCoercion</span></span><br><span data-ttu-id="c86cb-382">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-382">
         -TextBindings</span></span><br><span data-ttu-id="c86cb-383">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c86cb-383">
         -TextFile</span></span><br><span data-ttu-id="c86cb-384">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="c86cb-384">
         - Settings</span></span><br><span data-ttu-id="c86cb-385">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-385">
         -TextCoercion</span></span><br><span data-ttu-id="c86cb-386">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-386">
         -MatrixCoercion</span></span><br><span data-ttu-id="c86cb-387">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c86cb-387">
         - Matrix Bindings</span></span> </td> 
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="c86cb-388">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c86cb-388">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c86cb-389">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="c86cb-389">Platform</span></span></th>
    <th><span data-ttu-id="c86cb-390">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="c86cb-390">Extension points</span></span></th> 
    <th><span data-ttu-id="c86cb-391">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="c86cb-391">API requirement sets</span></span></th> 
    <th><span data-ttu-id="c86cb-392"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="c86cb-392"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-393">Office Online</span><span class="sxs-lookup"><span data-stu-id="c86cb-393">Office Online</span></span></td>
    <td> <span data-ttu-id="c86cb-394">- Contenu</span><span class="sxs-lookup"><span data-stu-id="c86cb-394">- Content</span></span><br><span data-ttu-id="c86cb-395">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="c86cb-395">
         - Taskpane</span></span><br><span data-ttu-id="c86cb-396">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-396">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c86cb-397">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-397">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c86cb-398">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c86cb-398">-ActiveView</span></span><br><span data-ttu-id="c86cb-399">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c86cb-399">
         -CompressedFile</span></span><br><span data-ttu-id="c86cb-400">
         - File</span><span class="sxs-lookup"><span data-stu-id="c86cb-400">
         - File</span></span><br><span data-ttu-id="c86cb-401">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="c86cb-401">
         - Selection</span></span><br><span data-ttu-id="c86cb-402">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="c86cb-402">
         - Settings</span></span><br><span data-ttu-id="c86cb-403">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-403">
         -TextCoercion</span></span><br><span data-ttu-id="c86cb-404">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-404">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-405">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="c86cb-405">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="c86cb-406">- Contenu</span><span class="sxs-lookup"><span data-stu-id="c86cb-406">- Content</span></span><br><span data-ttu-id="c86cb-407">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="c86cb-407">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="c86cb-408">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="c86cb-408">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="c86cb-409">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c86cb-409">-ActiveView</span></span><br><span data-ttu-id="c86cb-410">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c86cb-410">
         -CompressedFile</span></span><br><span data-ttu-id="c86cb-411">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-411">
         -DocumentEvents</span></span><br><span data-ttu-id="c86cb-412">
         - File</span><span class="sxs-lookup"><span data-stu-id="c86cb-412">
         - File</span></span><br><span data-ttu-id="c86cb-413">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="c86cb-413">
         - Selection</span></span><br><span data-ttu-id="c86cb-414">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="c86cb-414">
         - Settings</span></span><br><span data-ttu-id="c86cb-415">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-415">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-416">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="c86cb-416">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="c86cb-417">- Contenu</span><span class="sxs-lookup"><span data-stu-id="c86cb-417">- Content</span></span><br><span data-ttu-id="c86cb-418">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="c86cb-418">
         - Taskpane</span></span><br><span data-ttu-id="c86cb-419">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-419">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c86cb-420">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-420">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c86cb-421">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c86cb-421">-ActiveView</span></span><br><span data-ttu-id="c86cb-422">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c86cb-422">
         -CompressedFile</span></span><br><span data-ttu-id="c86cb-423">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-423">
         -DocumentEvents</span></span><br><span data-ttu-id="c86cb-424">
         - File</span><span class="sxs-lookup"><span data-stu-id="c86cb-424">
         - File</span></span><br><span data-ttu-id="c86cb-425">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="c86cb-425">
         - Selection</span></span><br><span data-ttu-id="c86cb-426">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="c86cb-426">
         - Settings</span></span><br><span data-ttu-id="c86cb-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-427">
         -TextCoercion</span></span><br><span data-ttu-id="c86cb-428">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-428">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-429">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="c86cb-429">Office for iOS</span></span></td>
    <td> <span data-ttu-id="c86cb-430">- Contenu</span><span class="sxs-lookup"><span data-stu-id="c86cb-430">- Content</span></span><br><span data-ttu-id="c86cb-431">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="c86cb-431">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="c86cb-432">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-432">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="c86cb-433">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c86cb-433">-ActiveView</span></span><br><span data-ttu-id="c86cb-434">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c86cb-434">
         -CompressedFile</span></span><br><span data-ttu-id="c86cb-435">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-435">
         -DocumentEvents</span></span><br><span data-ttu-id="c86cb-436">
         - File</span><span class="sxs-lookup"><span data-stu-id="c86cb-436">
         - File</span></span><br><span data-ttu-id="c86cb-437">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="c86cb-437">
         - Selection</span></span><br><span data-ttu-id="c86cb-438">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="c86cb-438">
         - Settings</span></span><br><span data-ttu-id="c86cb-439">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-439">
         -TextCoercion</span></span><br><span data-ttu-id="c86cb-440">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-440">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-441">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="c86cb-441">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="c86cb-442">- Contenu</span><span class="sxs-lookup"><span data-stu-id="c86cb-442">- Content</span></span><br><span data-ttu-id="c86cb-443">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="c86cb-443">
         - Taskpane</span></span><br><span data-ttu-id="c86cb-444">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-444">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c86cb-445">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-445">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c86cb-446">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c86cb-446">-ActiveView</span></span><br><span data-ttu-id="c86cb-447">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c86cb-447">
         -CompressedFile</span></span><br><span data-ttu-id="c86cb-448">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-448">
         -DocumentEvents</span></span><br><span data-ttu-id="c86cb-449">
         - File</span><span class="sxs-lookup"><span data-stu-id="c86cb-449">
         - File</span></span><br><span data-ttu-id="c86cb-450">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="c86cb-450">
         - Selection</span></span><br><span data-ttu-id="c86cb-451">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="c86cb-451">
         - Settings</span></span><br><span data-ttu-id="c86cb-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-452">
         -TextCoercion</span></span><br><span data-ttu-id="c86cb-453">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-453">
         -ImageCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="c86cb-454">OneNote</span><span class="sxs-lookup"><span data-stu-id="c86cb-454">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c86cb-455">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="c86cb-455">Platform</span></span></th>
    <th><span data-ttu-id="c86cb-456">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="c86cb-456">Extension points</span></span></th> 
    <th><span data-ttu-id="c86cb-457">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="c86cb-457">API requirement sets</span></span></th> 
    <th><span data-ttu-id="c86cb-458"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="c86cb-458"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="c86cb-459">Office Online</span><span class="sxs-lookup"><span data-stu-id="c86cb-459">Office Online</span></span></td>
    <td> <span data-ttu-id="c86cb-460">- Contenu</span><span class="sxs-lookup"><span data-stu-id="c86cb-460">- Content</span></span><br><span data-ttu-id="c86cb-461">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="c86cb-461">
         - Taskpane</span></span><br><span data-ttu-id="c86cb-462">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-462">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c86cb-463">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-463">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="c86cb-464">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c86cb-464">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c86cb-465">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c86cb-465">-DocumentEvents</span></span><br><span data-ttu-id="c86cb-466">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="c86cb-466">
         - Settings</span></span><br><span data-ttu-id="c86cb-467">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-467">
         -TextCoercion</span></span><br><span data-ttu-id="c86cb-468">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-468">
         -HtmlCoercion</span></span><br><span data-ttu-id="c86cb-469">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c86cb-469">
         -ImageCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="c86cb-470">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c86cb-470">See also</span></span>

- [<span data-ttu-id="c86cb-471">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="c86cb-471">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="c86cb-472">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="c86cb-472">Common API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="c86cb-473">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="c86cb-473">Add-in Commands requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="c86cb-474">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="c86cb-474">JavaScript API for Office reference</span></span>](https://dev.office.com/reference/add-ins/javascript-api-for-office)

