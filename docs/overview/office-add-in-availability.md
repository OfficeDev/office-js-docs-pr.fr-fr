---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint et OneNote.
ms.date: 08/30/2018
ms.openlocfilehash: 06fb073693bd8adca7d196f4361699ac3f54cee1
ms.sourcegitcommit: 78b28ae88d53bfef3134c09cc4336a5a8722c70b
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/01/2018
ms.locfileid: "23797300"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="4c624-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="4c624-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="4c624-104">Pour fonctionner comme prévu, il se peut que votre complément Office dépende d’un hôte Office spécifique, d’un ensemble de conditions requises, d’un membre d’API ou d’une version de l’API.</span><span class="sxs-lookup"><span data-stu-id="4c624-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="4c624-105">Les tableaux suivants contiennent la plateforme disponible, les points d’extension, les ensembles de conditions requises de l’API et les ensembles de conditions requises des API communes qui sont actuellement pris en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="4c624-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="4c624-106">Si une cellule de tableau contient un astérisque (\*), cela signifie que nous travaillons sur celle-ci.</span><span class="sxs-lookup"><span data-stu-id="4c624-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="4c624-107">Pour les ensembles de conditions requises pour Project ou Access, consultez les [ensembles de conditions requises communs à Office](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="4c624-107">For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="4c624-p103">Le numéro de build pour Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que les ensembles de conditions requises des API communes, ExcelApi 1.1 et WordApi 1.1.</span><span class="sxs-lookup"><span data-stu-id="4c624-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="4c624-110">Excel</span><span class="sxs-lookup"><span data-stu-id="4c624-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="4c624-111">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="4c624-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="4c624-112">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="4c624-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="4c624-113">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="4c624-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="4c624-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="4c624-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="4c624-115">Office Online</span></span></td>
    <td> <span data-ttu-id="4c624-116">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4c624-116">- Taskpane</span></span><br><span data-ttu-id="4c624-117">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="4c624-117">
        - Content</span></span><br><span data-ttu-id="4c624-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a>
    </span><span class="sxs-lookup"><span data-stu-id="4c624-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4c624-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4c624-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c624-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4c624-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c624-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4c624-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c624-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4c624-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c624-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4c624-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4c624-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4c624-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4c624-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="4c624-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4c624-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-127">
        -BindingEvents</span></span><br><span data-ttu-id="4c624-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c624-128">
        -CompressedFile</span></span><br><span data-ttu-id="4c624-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-129">
        -DocumentEvents</span></span><br><span data-ttu-id="4c624-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="4c624-130">
        - File</span></span><br><span data-ttu-id="4c624-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-131">
        -MatrixBindings</span></span><br><span data-ttu-id="4c624-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-132">
        -MatrixCoercion</span></span><br><span data-ttu-id="4c624-133">
        - Sélection</span><span class="sxs-lookup"><span data-stu-id="4c624-133">
        - Selection</span></span><br><span data-ttu-id="4c624-134">
        - Paramètres</span><span class="sxs-lookup"><span data-stu-id="4c624-134">
        - Settings</span></span><br><span data-ttu-id="4c624-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-135">
        -TableBindings</span></span><br><span data-ttu-id="4c624-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-136">
        -TableCoercion</span></span><br><span data-ttu-id="4c624-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-137">
        -TextBindings</span></span><br><span data-ttu-id="4c624-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-138">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-139">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4c624-139">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="4c624-140">
        - Volet Office</span><span class="sxs-lookup"><span data-stu-id="4c624-140">
        - Taskpane</span></span><br><span data-ttu-id="4c624-141">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="4c624-141">
        - Content</span></span></td>
    <td>  <span data-ttu-id="4c624-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4c624-143">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-143">
        -BindingEvents</span></span><br><span data-ttu-id="4c624-144">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c624-144">
        -CompressedFile</span></span><br><span data-ttu-id="4c624-145">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-145">
        -DocumentEvents</span></span><br><span data-ttu-id="4c624-146">
        - File</span><span class="sxs-lookup"><span data-stu-id="4c624-146">
        - File</span></span><br><span data-ttu-id="4c624-147">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-147">
        -ImageCoercion</span></span><br><span data-ttu-id="4c624-148">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-148">
        -MatrixBindings</span></span><br><span data-ttu-id="4c624-149">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-149">
        -MatrixCoercion</span></span><br><span data-ttu-id="4c624-150">
        - Sélection</span><span class="sxs-lookup"><span data-stu-id="4c624-150">
        - Selection</span></span><br><span data-ttu-id="4c624-151">
        - Paramètres</span><span class="sxs-lookup"><span data-stu-id="4c624-151">
        - Settings</span></span><br><span data-ttu-id="4c624-152">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-152">
        -TableBindings</span></span><br><span data-ttu-id="4c624-153">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-153">
        -TableCoercion</span></span><br><span data-ttu-id="4c624-154">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-154">
        -TextBindings</span></span><br><span data-ttu-id="4c624-155">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-155">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-156">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4c624-156">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="4c624-157">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4c624-157">- Taskpane</span></span><br><span data-ttu-id="4c624-158">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="4c624-158">
        - Content</span></span><br><span data-ttu-id="4c624-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4c624-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4c624-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4c624-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c624-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4c624-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c624-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4c624-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c624-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4c624-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c624-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4c624-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4c624-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4c624-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4c624-166">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="4c624-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4c624-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-168">-BindingEvents</span></span><br><span data-ttu-id="4c624-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c624-169">
        -CompressedFile</span></span><br><span data-ttu-id="4c624-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-170">
        -DocumentEvents</span></span><br><span data-ttu-id="4c624-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="4c624-171">
        - File</span></span><br><span data-ttu-id="4c624-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-172">
        -ImageCoercion</span></span><br><span data-ttu-id="4c624-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-173">
        -MatrixBindings</span></span><br><span data-ttu-id="4c624-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="4c624-175">
        - Sélection</span><span class="sxs-lookup"><span data-stu-id="4c624-175">
        - Selection</span></span><br><span data-ttu-id="4c624-176">
        - Paramètres</span><span class="sxs-lookup"><span data-stu-id="4c624-176">
        - Settings</span></span><br><span data-ttu-id="4c624-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-177">
        -TableBindings</span></span><br><span data-ttu-id="4c624-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-178">
        -TableCoercion</span></span><br><span data-ttu-id="4c624-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-179">
        -TextBindings</span></span><br><span data-ttu-id="4c624-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-181">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="4c624-181">Office for iOS</span></span></td>
    <td><span data-ttu-id="4c624-182">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4c624-182">- Taskpane</span></span><br><span data-ttu-id="4c624-183">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="4c624-183">
        - Content</span></span></td>
    <td><span data-ttu-id="4c624-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4c624-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c624-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4c624-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c624-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4c624-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c624-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4c624-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c624-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4c624-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4c624-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4c624-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4c624-190">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="4c624-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4c624-192">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-192">-BindingEvents</span></span><br><span data-ttu-id="4c624-193">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c624-193">
        -CompressedFile</span></span><br><span data-ttu-id="4c624-194">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-194">
        -DocumentEvents</span></span><br><span data-ttu-id="4c624-195">
        - File</span><span class="sxs-lookup"><span data-stu-id="4c624-195">
        - File</span></span><br><span data-ttu-id="4c624-196">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-196">
        -ImageCoercion</span></span><br><span data-ttu-id="4c624-197">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-197">
        -MatrixBindings</span></span><br><span data-ttu-id="4c624-198">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-198">
        -MatrixCoercion</span></span><br><span data-ttu-id="4c624-199">
        - Sélection</span><span class="sxs-lookup"><span data-stu-id="4c624-199">
        - Selection</span></span><br><span data-ttu-id="4c624-200">
        - Paramètres</span><span class="sxs-lookup"><span data-stu-id="4c624-200">
        - Settings</span></span><br><span data-ttu-id="4c624-201">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-201">
        -TableBindings</span></span><br><span data-ttu-id="4c624-202">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-202">
        -TableCoercion</span></span><br><span data-ttu-id="4c624-203">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-203">
        -TextBindings</span></span><br><span data-ttu-id="4c624-204">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-204">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-205">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="4c624-205">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="4c624-206">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4c624-206">- Taskpane</span></span><br><span data-ttu-id="4c624-207">
        - Contenu</span><span class="sxs-lookup"><span data-stu-id="4c624-207">
        - Content</span></span><br><span data-ttu-id="4c624-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4c624-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4c624-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4c624-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c624-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4c624-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c624-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4c624-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c624-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4c624-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c624-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4c624-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4c624-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4c624-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4c624-215">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="4c624-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4c624-217">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-217">-BindingEvents</span></span><br><span data-ttu-id="4c624-218">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c624-218">
        -CompressedFile</span></span><br><span data-ttu-id="4c624-219">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-219">
        -DocumentEvents</span></span><br><span data-ttu-id="4c624-220">
        - File</span><span class="sxs-lookup"><span data-stu-id="4c624-220">
        - File</span></span><br><span data-ttu-id="4c624-221">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-221">
        -ImageCoercion</span></span><br><span data-ttu-id="4c624-222">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-222">
        -MatrixBindings</span></span><br><span data-ttu-id="4c624-223">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-223">
        -MatrixCoercion</span></span><br><span data-ttu-id="4c624-224">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c624-224">
        -PdfFile</span></span><br><span data-ttu-id="4c624-225">
        - Sélection</span><span class="sxs-lookup"><span data-stu-id="4c624-225">
        - Selection</span></span><br><span data-ttu-id="4c624-226">
        - Paramètres</span><span class="sxs-lookup"><span data-stu-id="4c624-226">
        - Settings</span></span><br><span data-ttu-id="4c624-227">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-227">
        -TableBindings</span></span><br><span data-ttu-id="4c624-228">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-228">
        -TableCoercion</span></span><br><span data-ttu-id="4c624-229">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-229">
        -TextBindings</span></span><br><span data-ttu-id="4c624-230">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-230">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="4c624-231">Outlook</span><span class="sxs-lookup"><span data-stu-id="4c624-231">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4c624-232">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="4c624-232">Platform</span></span></th>
    <th><span data-ttu-id="4c624-233">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="4c624-233">Extension points</span></span></th>
    <th><span data-ttu-id="4c624-234">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="4c624-234">API requirement sets</span></span></th>
    <th><span data-ttu-id="4c624-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="4c624-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-236">Office Online</span><span class="sxs-lookup"><span data-stu-id="4c624-236">Office Online</span></span></td>
    <td> <span data-ttu-id="4c624-237">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="4c624-237">- Mail Read</span></span><br><span data-ttu-id="4c624-238">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="4c624-238">
      - Mail Compose</span></span><br><span data-ttu-id="4c624-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4c624-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c624-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4c624-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c624-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4c624-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c624-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4c624-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c624-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4c624-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c624-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4c624-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4c624-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4c624-246">Non disponible</span><span class="sxs-lookup"><span data-stu-id="4c624-246">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-247">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4c624-247">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4c624-248">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="4c624-248">- Mail Read</span></span><br><span data-ttu-id="4c624-249">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="4c624-249">
      - Mail Compose</span></span><br><span data-ttu-id="4c624-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4c624-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c624-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4c624-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c624-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4c624-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c624-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4c624-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c624-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="4c624-255">Non disponible</span><span class="sxs-lookup"><span data-stu-id="4c624-255">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-256">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4c624-256">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4c624-257">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="4c624-257">- Mail Read</span></span><br><span data-ttu-id="4c624-258">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="4c624-258">
      - Mail Compose</span></span><br><span data-ttu-id="4c624-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4c624-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4c624-260">
      - Modules</span><span class="sxs-lookup"><span data-stu-id="4c624-260">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4c624-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4c624-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c624-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4c624-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c624-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4c624-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c624-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4c624-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c624-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4c624-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4c624-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4c624-267">Non disponible</span><span class="sxs-lookup"><span data-stu-id="4c624-267">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-268">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="4c624-268">Office for iOS</span></span></td>
    <td> <span data-ttu-id="4c624-269">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="4c624-269">- Mail Read</span></span><br><span data-ttu-id="4c624-270">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4c624-270">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c624-271">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-271">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4c624-272">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c624-272">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4c624-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c624-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4c624-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c624-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4c624-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c624-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="4c624-276">Non disponible</span><span class="sxs-lookup"><span data-stu-id="4c624-276">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-277">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="4c624-277">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4c624-278">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="4c624-278">- Mail Read</span></span><br><span data-ttu-id="4c624-279">
      - Composition de message</span><span class="sxs-lookup"><span data-stu-id="4c624-279">
      - Mail Compose</span></span><br><span data-ttu-id="4c624-280">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4c624-280">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c624-281">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-281">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4c624-282">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c624-282">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4c624-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c624-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4c624-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c624-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4c624-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c624-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4c624-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4c624-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4c624-287">Non disponible</span><span class="sxs-lookup"><span data-stu-id="4c624-287">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-288">Office pour Android</span><span class="sxs-lookup"><span data-stu-id="4c624-288">Office for Android</span></span></td>
    <td> <span data-ttu-id="4c624-289">- Lecture de message</span><span class="sxs-lookup"><span data-stu-id="4c624-289">- Mail Read</span></span><br><span data-ttu-id="4c624-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4c624-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c624-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4c624-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c624-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4c624-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c624-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4c624-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c624-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4c624-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c624-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="4c624-296">Non disponible</span><span class="sxs-lookup"><span data-stu-id="4c624-296">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="4c624-297">Word</span><span class="sxs-lookup"><span data-stu-id="4c624-297">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4c624-298">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="4c624-298">Platform</span></span></th>
    <th><span data-ttu-id="4c624-299">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="4c624-299">Extension points</span></span></th>
    <th><span data-ttu-id="4c624-300">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="4c624-300">API requirement sets</span></span></th>
    <th><span data-ttu-id="4c624-301"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="4c624-301"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-302">Office Online</span><span class="sxs-lookup"><span data-stu-id="4c624-302">Office Online</span></span></td>
    <td> <span data-ttu-id="4c624-303">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4c624-303">- Taskpane</span></span><br><span data-ttu-id="4c624-304">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4c624-304">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c624-305">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-305">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4c624-306">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c624-306">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4c624-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c624-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4c624-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c624-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-309">-BindingEvents</span></span><br><span data-ttu-id="4c624-310">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4c624-310">
         -</span></span><br><span data-ttu-id="4c624-311">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-311">
         -DocumentEvents</span></span><br><span data-ttu-id="4c624-312">
         - File</span><span class="sxs-lookup"><span data-stu-id="4c624-312">
         - File</span></span><br><span data-ttu-id="4c624-313">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-313">
         -HtmlCoercion</span></span><br><span data-ttu-id="4c624-314">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-314">
         -ImageCoercion</span></span><br><span data-ttu-id="4c624-315">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-315">
         -MatrixBindings</span></span><br><span data-ttu-id="4c624-316">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-316">
         -MatrixCoercion</span></span><br><span data-ttu-id="4c624-317">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-317">
         -OoxmlCoercion</span></span><br><span data-ttu-id="4c624-318">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c624-318">
         -PdfFile</span></span><br><span data-ttu-id="4c624-319">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="4c624-319">
         - Selection</span></span><br><span data-ttu-id="4c624-320">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="4c624-320">
         - Settings</span></span><br><span data-ttu-id="4c624-321">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-321">
         -TableBindings</span></span><br><span data-ttu-id="4c624-322">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-322">
         -TableCoercion</span></span><br><span data-ttu-id="4c624-323">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-323">
         -TextBindings</span></span><br><span data-ttu-id="4c624-324">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-324">
         -TextCoercion</span></span><br><span data-ttu-id="4c624-325">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4c624-325">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-326">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4c624-326">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4c624-327">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4c624-327">- Taskpane</span></span></td>
    <td> <span data-ttu-id="4c624-328">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-328">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c624-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-329">-BindingEvents</span></span><br><span data-ttu-id="4c624-330">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c624-330">
         -CompressedFile</span></span><br><span data-ttu-id="4c624-331">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4c624-331">
         -</span></span><br><span data-ttu-id="4c624-332">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-332">
         -DocumentEvents</span></span><br><span data-ttu-id="4c624-333">
         - File</span><span class="sxs-lookup"><span data-stu-id="4c624-333">
         - File</span></span><br><span data-ttu-id="4c624-334">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-334">
         -HtmlCoercion</span></span><br><span data-ttu-id="4c624-335">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-335">
         -ImageCoercion</span></span><br><span data-ttu-id="4c624-336">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-336">
         -MatrixBindings</span></span><br><span data-ttu-id="4c624-337">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-337">
         -MatrixCoercion</span></span><br><span data-ttu-id="4c624-338">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-338">
         -OoxmlCoercion</span></span><br><span data-ttu-id="4c624-339">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c624-339">
         -PdfFile</span></span><br><span data-ttu-id="4c624-340">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="4c624-340">
         - Selection</span></span><br><span data-ttu-id="4c624-341">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="4c624-341">
         - Settings</span></span><br><span data-ttu-id="4c624-342">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-342">
         -TableBindings</span></span><br><span data-ttu-id="4c624-343">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-343">
         -TableCoercion</span></span><br><span data-ttu-id="4c624-344">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-344">
         -TextBindings</span></span><br><span data-ttu-id="4c624-345">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-345">
         -TextCoercion</span></span><br><span data-ttu-id="4c624-346">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4c624-346">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-347">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4c624-347">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4c624-348">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4c624-348">- Taskpane</span></span><br><span data-ttu-id="4c624-349">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4c624-349">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c624-350">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-350">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4c624-351">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c624-351">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4c624-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c624-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4c624-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c624-354">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-354">-BindingEvents</span></span><br><span data-ttu-id="4c624-355">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c624-355">
         -CompressedFile</span></span><br><span data-ttu-id="4c624-356">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4c624-356">
         -</span></span><br><span data-ttu-id="4c624-357">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-357">
         -DocumentEvents</span></span><br><span data-ttu-id="4c624-358">
         - File</span><span class="sxs-lookup"><span data-stu-id="4c624-358">
         - File</span></span><br><span data-ttu-id="4c624-359">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-359">
         -HtmlCoercion</span></span><br><span data-ttu-id="4c624-360">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-360">
         -ImageCoercion</span></span><br><span data-ttu-id="4c624-361">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-361">
         -MatrixBindings</span></span><br><span data-ttu-id="4c624-362">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-362">
         -MatrixCoercion</span></span><br><span data-ttu-id="4c624-363">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-363">
         -OoxmlCoercion</span></span><br><span data-ttu-id="4c624-364">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c624-364">
         -PdfFile</span></span><br><span data-ttu-id="4c624-365">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="4c624-365">
         - Selection</span></span><br><span data-ttu-id="4c624-366">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="4c624-366">
         - Settings</span></span><br><span data-ttu-id="4c624-367">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-367">
         -TableBindings</span></span><br><span data-ttu-id="4c624-368">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-368">
         -TableCoercion</span></span><br><span data-ttu-id="4c624-369">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-369">
         -TextBindings</span></span><br><span data-ttu-id="4c624-370">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-370">
         -TextCoercion</span></span><br><span data-ttu-id="4c624-371">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4c624-371">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-372">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="4c624-372">Office for iOS</span></span></td>
    <td> <span data-ttu-id="4c624-373">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4c624-373">- Taskpane</span></span></td>
    <td> <span data-ttu-id="4c624-374">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-374">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4c624-375">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c624-375">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4c624-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c624-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4c624-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4c624-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4c624-378">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-378">-BindingEvents</span></span><br><span data-ttu-id="4c624-379">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c624-379">
         -CompressedFile</span></span><br><span data-ttu-id="4c624-380">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4c624-380">
         -</span></span><br><span data-ttu-id="4c624-381">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-381">
         -DocumentEvents</span></span><br><span data-ttu-id="4c624-382">
         - File</span><span class="sxs-lookup"><span data-stu-id="4c624-382">
         - File</span></span><br><span data-ttu-id="4c624-383">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-383">
         -HtmlCoercion</span></span><br><span data-ttu-id="4c624-384">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-384">
         -ImageCoercion</span></span><br><span data-ttu-id="4c624-385">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-385">
         -MatrixBindings</span></span><br><span data-ttu-id="4c624-386">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-386">
         -MatrixCoercion</span></span><br><span data-ttu-id="4c624-387">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-387">
         -OoxmlCoercion</span></span><br><span data-ttu-id="4c624-388">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c624-388">
         -PdfFile</span></span><br><span data-ttu-id="4c624-389">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="4c624-389">
         - Selection</span></span><br><span data-ttu-id="4c624-390">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="4c624-390">
         - Settings</span></span><br><span data-ttu-id="4c624-391">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-391">
         -TableBindings</span></span><br><span data-ttu-id="4c624-392">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-392">
         -TableCoercion</span></span><br><span data-ttu-id="4c624-393">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-393">
         -TextBindings</span></span><br><span data-ttu-id="4c624-394">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-394">
         -TextCoercion</span></span><br><span data-ttu-id="4c624-395">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4c624-395">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-396">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="4c624-396">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4c624-397">- Volet Office</span><span class="sxs-lookup"><span data-stu-id="4c624-397">- Taskpane</span></span><br><span data-ttu-id="4c624-398">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4c624-398">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c624-399">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-399">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4c624-400">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c624-400">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4c624-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c624-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4c624-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4c624-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4c624-403">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-403">-BindingEvents</span></span><br><span data-ttu-id="4c624-404">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c624-404">
         -CompressedFile</span></span><br><span data-ttu-id="4c624-405">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4c624-405">
         -</span></span><br><span data-ttu-id="4c624-406">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-406">
         -DocumentEvents</span></span><br><span data-ttu-id="4c624-407">
         - File</span><span class="sxs-lookup"><span data-stu-id="4c624-407">
         - File</span></span><br><span data-ttu-id="4c624-408">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-408">
         -HtmlCoercion</span></span><br><span data-ttu-id="4c624-409">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-409">
         -ImageCoercion</span></span><br><span data-ttu-id="4c624-410">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-410">
         -MatrixBindings</span></span><br><span data-ttu-id="4c624-411">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-411">
         -MatrixCoercion</span></span><br><span data-ttu-id="4c624-412">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-412">
         -OoxmlCoercion</span></span><br><span data-ttu-id="4c624-413">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c624-413">
         -PdfFile</span></span><br><span data-ttu-id="4c624-414">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="4c624-414">
         - Selection</span></span><br><span data-ttu-id="4c624-415">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="4c624-415">
         - Settings</span></span><br><span data-ttu-id="4c624-416">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-416">
         -TableBindings</span></span><br><span data-ttu-id="4c624-417">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-417">
         -TableCoercion</span></span><br><span data-ttu-id="4c624-418">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c624-418">
         -TextBindings</span></span><br><span data-ttu-id="4c624-419">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-419">
         -TextCoercion</span></span><br><span data-ttu-id="4c624-420">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4c624-420">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="4c624-421">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="4c624-421">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4c624-422">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="4c624-422">Platform</span></span></th>
    <th><span data-ttu-id="4c624-423">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="4c624-423">Extension points</span></span></th>
    <th><span data-ttu-id="4c624-424">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="4c624-424">API requirement sets</span></span></th>
    <th><span data-ttu-id="4c624-425"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="4c624-425"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-426">Office Online</span><span class="sxs-lookup"><span data-stu-id="4c624-426">Office Online</span></span></td>
    <td> <span data-ttu-id="4c624-427">- Contenu</span><span class="sxs-lookup"><span data-stu-id="4c624-427">- Content</span></span><br><span data-ttu-id="4c624-428">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="4c624-428">
         - Taskpane</span></span><br><span data-ttu-id="4c624-429">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4c624-429">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c624-430">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-430">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c624-431">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4c624-431">-ActiveView</span></span><br><span data-ttu-id="4c624-432">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c624-432">
         -CompressedFile</span></span><br><span data-ttu-id="4c624-433">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-433">
         -DocumentEvents</span></span><br><span data-ttu-id="4c624-434">
         - File</span><span class="sxs-lookup"><span data-stu-id="4c624-434">
         - File</span></span><br><span data-ttu-id="4c624-435">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-435">
         -ImageCoercion</span></span><br><span data-ttu-id="4c624-436">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c624-436">
         -PdfFile</span></span><br><span data-ttu-id="4c624-437">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="4c624-437">
         - Selection</span></span><br><span data-ttu-id="4c624-438">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="4c624-438">
         - Settings</span></span><br><span data-ttu-id="4c624-439">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-439">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-440">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4c624-440">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4c624-441">- Contenu</span><span class="sxs-lookup"><span data-stu-id="4c624-441">- Content</span></span><br><span data-ttu-id="4c624-442">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="4c624-442">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="4c624-443">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4c624-443">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4c624-444">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4c624-444">-ActiveView</span></span><br><span data-ttu-id="4c624-445">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c624-445">
         -CompressedFile</span></span><br><span data-ttu-id="4c624-446">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-446">
         -DocumentEvents</span></span><br><span data-ttu-id="4c624-447">
         - File</span><span class="sxs-lookup"><span data-stu-id="4c624-447">
         - File</span></span><br><span data-ttu-id="4c624-448">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-448">
         -ImageCoercion</span></span><br><span data-ttu-id="4c624-449">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c624-449">
         -PdfFile</span></span><br><span data-ttu-id="4c624-450">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="4c624-450">
         - Selection</span></span><br><span data-ttu-id="4c624-451">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="4c624-451">
         - Settings</span></span><br><span data-ttu-id="4c624-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-452">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-453">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="4c624-453">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4c624-454">- Contenu</span><span class="sxs-lookup"><span data-stu-id="4c624-454">- Content</span></span><br><span data-ttu-id="4c624-455">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="4c624-455">
         - Taskpane</span></span><br><span data-ttu-id="4c624-456">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4c624-456">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c624-457">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-457">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c624-458">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4c624-458">-ActiveView</span></span><br><span data-ttu-id="4c624-459">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c624-459">
         -CompressedFile</span></span><br><span data-ttu-id="4c624-460">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-460">
         -DocumentEvents</span></span><br><span data-ttu-id="4c624-461">
         - File</span><span class="sxs-lookup"><span data-stu-id="4c624-461">
         - File</span></span><br><span data-ttu-id="4c624-462">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-462">
         -ImageCoercion</span></span><br><span data-ttu-id="4c624-463">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c624-463">
         -PdfFile</span></span><br><span data-ttu-id="4c624-464">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="4c624-464">
         - Selection</span></span><br><span data-ttu-id="4c624-465">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="4c624-465">
         - Settings</span></span><br><span data-ttu-id="4c624-466">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-466">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-467">Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="4c624-467">Office for iOS</span></span></td>
    <td> <span data-ttu-id="4c624-468">- Contenu</span><span class="sxs-lookup"><span data-stu-id="4c624-468">- Content</span></span><br><span data-ttu-id="4c624-469">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="4c624-469">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="4c624-470">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-470">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="4c624-471">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4c624-471">-ActiveView</span></span><br><span data-ttu-id="4c624-472">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c624-472">
         -CompressedFile</span></span><br><span data-ttu-id="4c624-473">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-473">
         -DocumentEvents</span></span><br><span data-ttu-id="4c624-474">
         - File</span><span class="sxs-lookup"><span data-stu-id="4c624-474">
         - File</span></span><br><span data-ttu-id="4c624-475">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c624-475">
         -PdfFile</span></span><br><span data-ttu-id="4c624-476">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="4c624-476">
         - Selection</span></span><br><span data-ttu-id="4c624-477">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="4c624-477">
         - Settings</span></span><br><span data-ttu-id="4c624-478">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-478">
         -TextCoercion</span></span><br><span data-ttu-id="4c624-479">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-479">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-480">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="4c624-480">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4c624-481">- Contenu</span><span class="sxs-lookup"><span data-stu-id="4c624-481">- Content</span></span><br><span data-ttu-id="4c624-482">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="4c624-482">
         - Taskpane</span></span><br><span data-ttu-id="4c624-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4c624-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c624-484">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-484">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c624-485">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4c624-485">-ActiveView</span></span><br><span data-ttu-id="4c624-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c624-486">
         -CompressedFile</span></span><br><span data-ttu-id="4c624-487">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-487">
         -DocumentEvents</span></span><br><span data-ttu-id="4c624-488">
         - File</span><span class="sxs-lookup"><span data-stu-id="4c624-488">
         - File</span></span><br><span data-ttu-id="4c624-489">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-489">
         -ImageCoercion</span></span><br><span data-ttu-id="4c624-490">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c624-490">
         -PdfFile</span></span><br><span data-ttu-id="4c624-491">
         - Sélection</span><span class="sxs-lookup"><span data-stu-id="4c624-491">
         - Selection</span></span><br><span data-ttu-id="4c624-492">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="4c624-492">
         - Settings</span></span><br><span data-ttu-id="4c624-493">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-493">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="4c624-494">OneNote</span><span class="sxs-lookup"><span data-stu-id="4c624-494">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4c624-495">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="4c624-495">Platform</span></span></th>
    <th><span data-ttu-id="4c624-496">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="4c624-496">Extension points</span></span></th>
    <th><span data-ttu-id="4c624-497">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="4c624-497">API requirement sets</span></span></th>
    <th><span data-ttu-id="4c624-498"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>API communes</b></a></span><span class="sxs-lookup"><span data-stu-id="4c624-498"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="4c624-499">Office Online</span><span class="sxs-lookup"><span data-stu-id="4c624-499">Office Online</span></span></td>
    <td> <span data-ttu-id="4c624-500">- Contenu</span><span class="sxs-lookup"><span data-stu-id="4c624-500">- Content</span></span><br><span data-ttu-id="4c624-501">
         - Volet Office</span><span class="sxs-lookup"><span data-stu-id="4c624-501">
         - Taskpane</span></span><br><span data-ttu-id="4c624-502">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Commandes de complément</a></span><span class="sxs-lookup"><span data-stu-id="4c624-502">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c624-503">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-503">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="4c624-504">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c624-504">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c624-505">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c624-505">-DocumentEvents</span></span><br><span data-ttu-id="4c624-506">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-506">
         -HtmlCoercion</span></span><br><span data-ttu-id="4c624-507">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-507">
         -ImageCoercion</span></span><br><span data-ttu-id="4c624-508">
         - Paramètres</span><span class="sxs-lookup"><span data-stu-id="4c624-508">
         - Settings</span></span><br><span data-ttu-id="4c624-509">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c624-509">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="4c624-510">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="4c624-510">See also</span></span>

- [<span data-ttu-id="4c624-511">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="4c624-511">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="4c624-512">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="4c624-512">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="4c624-513">Ensembles de conditions requises concernant les commandes de complément</span><span class="sxs-lookup"><span data-stu-id="4c624-513">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="4c624-514">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="4c624-514">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)
