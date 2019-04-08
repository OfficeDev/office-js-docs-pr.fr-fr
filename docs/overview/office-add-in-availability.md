---
title: Disponibilité des compléments Office sur les plateformes et les hôtes
description: Ensembles de conditions requises pris en charge pour Excel, Word, Outlook, PowerPoint, OneNote et Project.
ms.date: 04/03/2019
localization_priority: Priority
ms.openlocfilehash: a9ecd44edf9221a403eb42756cd1e9f5e676ad01
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/04/2019
ms.locfileid: "31477592"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="da394-103">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="da394-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="da394-p101">Pour fonctionner comme prévu, votre complément Office peut dépendre d'un hôte Office spécifique, d'un ensemble de conditions requises, d'un membre API ou d'une version de l'API. Les tableaux suivants contiennent les plates-formes disponibles, les points d'extension, les ensembles de conditions requises de l’API et les API communes qui sont actuellement prises en charge pour chaque application Office.</span><span class="sxs-lookup"><span data-stu-id="da394-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="da394-106">La version initiale d’Office 2016 installée via MSI contient uniquement les ensembles de conditions ExcelApi 1.1, WordApi 1.1 et API commune.</span><span class="sxs-lookup"><span data-stu-id="da394-106">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="da394-107">Pour plus d’informations sur l’historique de mise à jour des différentes versions d’Office, consultez la section [Voir aussi](#see-also).</span><span class="sxs-lookup"><span data-stu-id="da394-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="da394-108">Excel</span><span class="sxs-lookup"><span data-stu-id="da394-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="da394-109">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="da394-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="da394-110">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="da394-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="da394-111">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="da394-111">API requirement sets</span></span></th>
    <th style="width:40%"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="da394-112">API communes</span><span class="sxs-lookup"><span data-stu-id="da394-112">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="da394-113">Office Online</span></span></td>
    <td> - <span data-ttu-id="da394-114">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-114">TaskPane</span></span><br>
        - <span data-ttu-id="da394-115">Content</span><span class="sxs-lookup"><span data-stu-id="da394-115">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-116">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-116">add-in commands</span></span></a>
    </td>
    <td>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-117">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-117">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-118">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-118">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-119">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-119">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-120">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="da394-120">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-121">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="da394-121">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-122">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="da394-122">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-123">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="da394-123">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-124">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="da394-124">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-125">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-125">DialogApi 1.1</span></span></a></td>
    <td>
        - <span data-ttu-id="da394-126">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da394-126">BindingEvents</span></span><br>
        - <span data-ttu-id="da394-127">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-127">CompressedFile</span></span><br>
        - <span data-ttu-id="da394-128">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-128">DocumentEvents</span></span><br>
        - <span data-ttu-id="da394-129">File</span><span class="sxs-lookup"><span data-stu-id="da394-129">File</span></span><br>
        - <span data-ttu-id="da394-130">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da394-130">MatrixBindings</span></span><br>
        - <span data-ttu-id="da394-131">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-131">MatrixCoercion</span></span><br>
        - <span data-ttu-id="da394-132">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-132">Selection</span></span><br>
        - <span data-ttu-id="da394-133">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-133">Settings</span></span><br>
        - <span data-ttu-id="da394-134">TableBindings</span><span class="sxs-lookup"><span data-stu-id="da394-134">TableBindings</span></span><br>
        - <span data-ttu-id="da394-135">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-135">TableCoercion</span></span><br>
        - <span data-ttu-id="da394-136">TextBindings</span><span class="sxs-lookup"><span data-stu-id="da394-136">TextBindings</span></span><br>
        - <span data-ttu-id="da394-137">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-137">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-138">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="da394-138">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="da394-139">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-139">TaskPane</span></span><br>
        - <span data-ttu-id="da394-140">Content</span><span class="sxs-lookup"><span data-stu-id="da394-140">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-141">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-141">add-in commands</span></span></a>
    </td>
    <td>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-142">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-142">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-143">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-143">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-144">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-144">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-145">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="da394-145">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-146">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="da394-146">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-147">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="da394-147">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-148">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="da394-148">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-149">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="da394-149">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-150">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-150">DialogApi 1.1</span></span></a></td>
    <td>
        - <span data-ttu-id="da394-151">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da394-151">BindingEvents</span></span><br>
        - <span data-ttu-id="da394-152">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-152">CompressedFile</span></span><br>
        - <span data-ttu-id="da394-153">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-153">DocumentEvents</span></span><br>
        - <span data-ttu-id="da394-154">File</span><span class="sxs-lookup"><span data-stu-id="da394-154">File</span></span><br>
        - <span data-ttu-id="da394-155">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da394-155">MatrixBindings</span></span><br>
        - <span data-ttu-id="da394-156">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-156">MatrixCoercion</span></span><br>
        - <span data-ttu-id="da394-157">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-157">Selection</span></span><br>
        - <span data-ttu-id="da394-158">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-158">Settings</span></span><br>
        - <span data-ttu-id="da394-159">TableBindings</span><span class="sxs-lookup"><span data-stu-id="da394-159">TableBindings</span></span><br>
        - <span data-ttu-id="da394-160">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-160">TableCoercion</span></span><br>
        - <span data-ttu-id="da394-161">TextBindings</span><span class="sxs-lookup"><span data-stu-id="da394-161">TextBindings</span></span><br>
        - <span data-ttu-id="da394-162">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-162">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-163">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="da394-163">Office 2019 for Windows</span></span></td>
    <td>- <span data-ttu-id="da394-164">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-164">TaskPane</span></span><br>
        - <span data-ttu-id="da394-165">Content</span><span class="sxs-lookup"><span data-stu-id="da394-165">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-166">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-166">add-in commands</span></span></a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-167">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-167">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-168">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-168">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-169">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-169">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-170">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="da394-170">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-171">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="da394-171">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-172">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="da394-172">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-173">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="da394-173">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-174">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="da394-174">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-175">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-175">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="da394-176">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da394-176">BindingEvents</span></span><br>
        - <span data-ttu-id="da394-177">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-177">CompressedFile</span></span><br>
        - <span data-ttu-id="da394-178">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-178">DocumentEvents</span></span><br>
        - <span data-ttu-id="da394-179">File</span><span class="sxs-lookup"><span data-stu-id="da394-179">File</span></span><br>
        - <span data-ttu-id="da394-180">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-180">ImageCoercion</span></span><br>
        - <span data-ttu-id="da394-181">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da394-181">MatrixBindings</span></span><br>
        - <span data-ttu-id="da394-182">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-182">MatrixCoercion</span></span><br>
        - <span data-ttu-id="da394-183">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-183">Selection</span></span><br>
        - <span data-ttu-id="da394-184">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-184">Settings</span></span><br>
        - <span data-ttu-id="da394-185">TableBindings</span><span class="sxs-lookup"><span data-stu-id="da394-185">TableBindings</span></span><br>
        - <span data-ttu-id="da394-186">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-186">TableCoercion</span></span><br>
        - <span data-ttu-id="da394-187">TextBindings</span><span class="sxs-lookup"><span data-stu-id="da394-187">TextBindings</span></span><br>
        - <span data-ttu-id="da394-188">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-188">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-189">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="da394-189">Office 2016 for Windows</span></span></td>
    <td>- <span data-ttu-id="da394-190">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-190">TaskPane</span></span><br>
        - <span data-ttu-id="da394-191">Content</span><span class="sxs-lookup"><span data-stu-id="da394-191">Content</span></span></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-192">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-192">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-193">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-193">DialogApi 1.1</span></span></a>*</td>
    <td>- <span data-ttu-id="da394-194">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da394-194">BindingEvents</span></span><br>
        - <span data-ttu-id="da394-195">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-195">CompressedFile</span></span><br>
        - <span data-ttu-id="da394-196">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-196">DocumentEvents</span></span><br>
        - <span data-ttu-id="da394-197">File</span><span class="sxs-lookup"><span data-stu-id="da394-197">File</span></span><br>
        - <span data-ttu-id="da394-198">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-198">ImageCoercion</span></span><br>
        - <span data-ttu-id="da394-199">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da394-199">MatrixBindings</span></span><br>
        - <span data-ttu-id="da394-200">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-200">MatrixCoercion</span></span><br>
        - <span data-ttu-id="da394-201">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-201">Selection</span></span><br>
        - <span data-ttu-id="da394-202">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-202">Settings</span></span><br>
        - <span data-ttu-id="da394-203">TableBindings</span><span class="sxs-lookup"><span data-stu-id="da394-203">TableBindings</span></span><br>
        - <span data-ttu-id="da394-204">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-204">TableCoercion</span></span><br>
        - <span data-ttu-id="da394-205">TextBindings</span><span class="sxs-lookup"><span data-stu-id="da394-205">TextBindings</span></span><br>
        - <span data-ttu-id="da394-206">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-206">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-207">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="da394-207">Office 2013 for Windows</span></span></td>
    <td>
        - <span data-ttu-id="da394-208">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-208">TaskPane</span></span><br>
        - <span data-ttu-id="da394-209">Content</span><span class="sxs-lookup"><span data-stu-id="da394-209">Content</span></span></td>
    <td>  - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-210">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-210">DialogApi 1.1</span></span></a>*</td>
    <td>
        - <span data-ttu-id="da394-211">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da394-211">BindingEvents</span></span><br>
        - <span data-ttu-id="da394-212">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-212">CompressedFile</span></span><br>
        - <span data-ttu-id="da394-213">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-213">DocumentEvents</span></span><br>
        - <span data-ttu-id="da394-214">File</span><span class="sxs-lookup"><span data-stu-id="da394-214">File</span></span><br>
        - <span data-ttu-id="da394-215">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-215">ImageCoercion</span></span><br>
        - <span data-ttu-id="da394-216">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da394-216">MatrixBindings</span></span><br>
        - <span data-ttu-id="da394-217">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-217">MatrixCoercion</span></span><br>
        - <span data-ttu-id="da394-218">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-218">Selection</span></span><br>
        - <span data-ttu-id="da394-219">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-219">Settings</span></span><br>
        - <span data-ttu-id="da394-220">TableBindings</span><span class="sxs-lookup"><span data-stu-id="da394-220">TableBindings</span></span><br>
        - <span data-ttu-id="da394-221">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-221">TableCoercion</span></span><br>
        - <span data-ttu-id="da394-222">TextBindings</span><span class="sxs-lookup"><span data-stu-id="da394-222">TextBindings</span></span><br>
        - <span data-ttu-id="da394-223">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-223">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-224">Office 365 pour iPad</span><span class="sxs-lookup"><span data-stu-id="da394-224">Office 365 for iPad</span></span></td>
    <td>- <span data-ttu-id="da394-225">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-225">TaskPane</span></span><br>
        - <span data-ttu-id="da394-226">Content</span><span class="sxs-lookup"><span data-stu-id="da394-226">Content</span></span></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-227">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-227">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-228">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-228">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-229">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-229">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-230">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="da394-230">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-231">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="da394-231">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-232">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="da394-232">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-233">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="da394-233">ExcelApi 1.7</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-234">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="da394-234">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-235">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-235">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="da394-236">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da394-236">BindingEvents</span></span><br>
        - <span data-ttu-id="da394-237">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-237">CompressedFile</span></span><br>
        - <span data-ttu-id="da394-238">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-238">DocumentEvents</span></span><br>
        - <span data-ttu-id="da394-239">File</span><span class="sxs-lookup"><span data-stu-id="da394-239">File</span></span><br>
        - <span data-ttu-id="da394-240">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-240">ImageCoercion</span></span><br>
        - <span data-ttu-id="da394-241">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da394-241">MatrixBindings</span></span><br>
        - <span data-ttu-id="da394-242">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-242">MatrixCoercion</span></span><br>
        - <span data-ttu-id="da394-243">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-243">Selection</span></span><br>
        - <span data-ttu-id="da394-244">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-244">Settings</span></span><br>
        - <span data-ttu-id="da394-245">TableBindings</span><span class="sxs-lookup"><span data-stu-id="da394-245">TableBindings</span></span><br>
        - <span data-ttu-id="da394-246">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-246">TableCoercion</span></span><br>
        - <span data-ttu-id="da394-247">TextBindings</span><span class="sxs-lookup"><span data-stu-id="da394-247">TextBindings</span></span><br>
        - <span data-ttu-id="da394-248">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-248">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-249">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="da394-249">Office 365 for Mac</span></span></td>
    <td>- <span data-ttu-id="da394-250">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-250">TaskPane</span></span><br>
        - <span data-ttu-id="da394-251">Content</span><span class="sxs-lookup"><span data-stu-id="da394-251">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-252">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-252">add-in commands</span></span></a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-253">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-253">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-254">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-254">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-255">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-255">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-256">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="da394-256">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-257">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="da394-257">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-258">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="da394-258">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-259">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="da394-259">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-260">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="da394-260">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-261">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-261">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="da394-262">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da394-262">BindingEvents</span></span><br>
        - <span data-ttu-id="da394-263">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-263">CompressedFile</span></span><br>
        - <span data-ttu-id="da394-264">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-264">DocumentEvents</span></span><br>
        - <span data-ttu-id="da394-265">File</span><span class="sxs-lookup"><span data-stu-id="da394-265">File</span></span><br>
        - <span data-ttu-id="da394-266">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-266">ImageCoercion</span></span><br>
        - <span data-ttu-id="da394-267">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da394-267">MatrixBindings</span></span><br>
        - <span data-ttu-id="da394-268">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-268">MatrixCoercion</span></span><br>
        - <span data-ttu-id="da394-269">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-269">PdfFile</span></span><br>
        - <span data-ttu-id="da394-270">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-270">Selection</span></span><br>
        - <span data-ttu-id="da394-271">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-271">Settings</span></span><br>
        - <span data-ttu-id="da394-272">TableBindings</span><span class="sxs-lookup"><span data-stu-id="da394-272">TableBindings</span></span><br>
        - <span data-ttu-id="da394-273">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-273">TableCoercion</span></span><br>
        - <span data-ttu-id="da394-274">TextBindings</span><span class="sxs-lookup"><span data-stu-id="da394-274">TextBindings</span></span><br>
        - <span data-ttu-id="da394-275">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-275">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-276">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="da394-276">Office 2019 for Mac</span></span></td>
    <td>- <span data-ttu-id="da394-277">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-277">TaskPane</span></span><br>
        - <span data-ttu-id="da394-278">Content</span><span class="sxs-lookup"><span data-stu-id="da394-278">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-279">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-279">add-in commands</span></span></a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-280">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-280">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-281">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-281">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-282">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-282">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-283">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="da394-283">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-284">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="da394-284">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-285">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="da394-285">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-286">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="da394-286">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-287">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="da394-287">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-288">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-288">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="da394-289">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da394-289">BindingEvents</span></span><br>
        - <span data-ttu-id="da394-290">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-290">CompressedFile</span></span><br>
        - <span data-ttu-id="da394-291">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-291">DocumentEvents</span></span><br>
        - <span data-ttu-id="da394-292">File</span><span class="sxs-lookup"><span data-stu-id="da394-292">File</span></span><br>
        - <span data-ttu-id="da394-293">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-293">ImageCoercion</span></span><br>
        - <span data-ttu-id="da394-294">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da394-294">MatrixBindings</span></span><br>
        - <span data-ttu-id="da394-295">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-295">MatrixCoercion</span></span><br>
        - <span data-ttu-id="da394-296">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-296">PdfFile</span></span><br>
        - <span data-ttu-id="da394-297">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-297">Selection</span></span><br>
        - <span data-ttu-id="da394-298">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-298">Settings</span></span><br>
        - <span data-ttu-id="da394-299">TableBindings</span><span class="sxs-lookup"><span data-stu-id="da394-299">TableBindings</span></span><br>
        - <span data-ttu-id="da394-300">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-300">TableCoercion</span></span><br>
        - <span data-ttu-id="da394-301">TextBindings</span><span class="sxs-lookup"><span data-stu-id="da394-301">TextBindings</span></span><br>
        - <span data-ttu-id="da394-302">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-302">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-303">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="da394-303">Office 2016 for Mac</span></span></td>
    <td>- <span data-ttu-id="da394-304">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-304">TaskPane</span></span><br>
        - <span data-ttu-id="da394-305">Content</span><span class="sxs-lookup"><span data-stu-id="da394-305">Content</span></span></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="da394-306">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-306">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-307">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-307">DialogApi 1.1</span></span></a>*</td>
    <td>- <span data-ttu-id="da394-308">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da394-308">BindingEvents</span></span><br>
        - <span data-ttu-id="da394-309">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-309">CompressedFile</span></span><br>
        - <span data-ttu-id="da394-310">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-310">DocumentEvents</span></span><br>
        - <span data-ttu-id="da394-311">File</span><span class="sxs-lookup"><span data-stu-id="da394-311">File</span></span><br>
        - <span data-ttu-id="da394-312">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-312">ImageCoercion</span></span><br>
        - <span data-ttu-id="da394-313">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da394-313">MatrixBindings</span></span><br>
        - <span data-ttu-id="da394-314">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-314">MatrixCoercion</span></span><br>
        - <span data-ttu-id="da394-315">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-315">PdfFile</span></span><br>
        - <span data-ttu-id="da394-316">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-316">Selection</span></span><br>
        - <span data-ttu-id="da394-317">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-317">Settings</span></span><br>
        - <span data-ttu-id="da394-318">TableBindings</span><span class="sxs-lookup"><span data-stu-id="da394-318">TableBindings</span></span><br>
        - <span data-ttu-id="da394-319">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-319">TableCoercion</span></span><br>
        - <span data-ttu-id="da394-320">TextBindings</span><span class="sxs-lookup"><span data-stu-id="da394-320">TextBindings</span></span><br>
        - <span data-ttu-id="da394-321">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-321">TextCoercion</span></span></td>
  </tr>
</table>

*<span data-ttu-id="da394-322">&ast; : ajouté avec les mises à jour postérieures à la publication.</span><span class="sxs-lookup"><span data-stu-id="da394-322">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="outlook"></a><span data-ttu-id="da394-323">Outlook</span><span class="sxs-lookup"><span data-stu-id="da394-323">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="da394-324">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="da394-324">Platform</span></span></th>
    <th><span data-ttu-id="da394-325">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="da394-325">Extension points</span></span></th>
    <th><span data-ttu-id="da394-326">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="da394-326">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="da394-327">API communes</span><span class="sxs-lookup"><span data-stu-id="da394-327">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-328">Office Online</span><span class="sxs-lookup"><span data-stu-id="da394-328">Office Online</span></span></td>
    <td> - <span data-ttu-id="da394-329">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="da394-329">Mail Read</span></span><br>
      - <span data-ttu-id="da394-330">Composition de message</span><span class="sxs-lookup"><span data-stu-id="da394-330">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-331">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-331">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="da394-332">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-332">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="da394-333">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-333">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="da394-334">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-334">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="da394-335">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="da394-335">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="da394-336">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="da394-336">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="da394-337">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="da394-337">Mailbox 1.6</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7"><span data-ttu-id="da394-338">Mailbox 1.7</span><span class="sxs-lookup"><span data-stu-id="da394-338">Mailbox 1.7</span></span></a></td>
    <td><span data-ttu-id="da394-339">Non disponible</span><span class="sxs-lookup"><span data-stu-id="da394-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-340">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="da394-340">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="da394-341">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="da394-341">Mail Read</span></span><br>
      - <span data-ttu-id="da394-342">Composition de message</span><span class="sxs-lookup"><span data-stu-id="da394-342">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-343">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-343">add-in commands</span></span></a><br>
      - <span data-ttu-id="da394-344">Modules</span><span class="sxs-lookup"><span data-stu-id="da394-344">Modules</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="da394-345">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-345">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="da394-346">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-346">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="da394-347">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-347">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="da394-348">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="da394-348">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="da394-349">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="da394-349">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="da394-350">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="da394-350">Mailbox 1.6</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7"><span data-ttu-id="da394-351">Mailbox 1.7</span><span class="sxs-lookup"><span data-stu-id="da394-351">Mailbox 1.7</span></span></a></td>
    <td><span data-ttu-id="da394-352">Non disponible</span><span class="sxs-lookup"><span data-stu-id="da394-352">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-353">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="da394-353">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="da394-354">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="da394-354">Mail Read</span></span><br>
      - <span data-ttu-id="da394-355">Composition de message</span><span class="sxs-lookup"><span data-stu-id="da394-355">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-356">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-356">add-in commands</span></span></a><br>
      - <span data-ttu-id="da394-357">Modules</span><span class="sxs-lookup"><span data-stu-id="da394-357">Modules</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="da394-358">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-358">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="da394-359">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-359">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="da394-360">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-360">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="da394-361">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="da394-361">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="da394-362">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="da394-362">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="da394-363">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="da394-363">Mailbox 1.6</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7"><span data-ttu-id="da394-364">Mailbox 1.7</span><span class="sxs-lookup"><span data-stu-id="da394-364">Mailbox 1.7</span></span></a></td>
    <td><span data-ttu-id="da394-365">Non disponible</span><span class="sxs-lookup"><span data-stu-id="da394-365">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-366">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="da394-366">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="da394-367">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="da394-367">Mail Read</span></span><br>
      - <span data-ttu-id="da394-368">Composition de message</span><span class="sxs-lookup"><span data-stu-id="da394-368">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-369">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-369">add-in commands</span></span></a><br>
      - <span data-ttu-id="da394-370">Modules</span><span class="sxs-lookup"><span data-stu-id="da394-370">Modules</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="da394-371">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-371">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="da394-372">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-372">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="da394-373">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-373">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="da394-374">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="da394-374">Mailbox 1.4</span></span></a>*</td>
    <td><span data-ttu-id="da394-375">Non disponible</span><span class="sxs-lookup"><span data-stu-id="da394-375">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-376">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="da394-376">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="da394-377">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="da394-377">Mail Read</span></span><br>
      - <span data-ttu-id="da394-378">Composition de message</span><span class="sxs-lookup"><span data-stu-id="da394-378">Mail Compose</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="da394-379">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-379">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="da394-380">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-380">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="da394-381">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-381">Mailbox 1.3</span></span></a>*<br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="da394-382">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="da394-382">Mailbox 1.4</span></span></a>*</td>
    <td><span data-ttu-id="da394-383">Non disponible</span><span class="sxs-lookup"><span data-stu-id="da394-383">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-384">Office 365 pour iOS</span><span class="sxs-lookup"><span data-stu-id="da394-384">Office 365 for iOS</span></span></td>
    <td> - <span data-ttu-id="da394-385">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="da394-385">Mail Read</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-386">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-386">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="da394-387">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-387">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="da394-388">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-388">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="da394-389">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-389">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="da394-390">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="da394-390">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="da394-391">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="da394-391">Mailbox 1.5</span></span></a></td>
    <td><span data-ttu-id="da394-392">Non disponible</span><span class="sxs-lookup"><span data-stu-id="da394-392">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-393">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="da394-393">Office 365 for Mac</span></span></td>
    <td> - <span data-ttu-id="da394-394">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="da394-394">Mail Read</span></span><br>
      - <span data-ttu-id="da394-395">Composition de message</span><span class="sxs-lookup"><span data-stu-id="da394-395">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-396">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-396">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="da394-397">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-397">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="da394-398">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-398">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="da394-399">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-399">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="da394-400">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="da394-400">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="da394-401">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="da394-401">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="da394-402">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="da394-402">Mailbox 1.6</span></span></a></td>
    <td><span data-ttu-id="da394-403">Non disponible</span><span class="sxs-lookup"><span data-stu-id="da394-403">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-404">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="da394-404">Office 2019 for Mac</span></span></td>
    <td> - <span data-ttu-id="da394-405">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="da394-405">Mail Read</span></span><br>
      - <span data-ttu-id="da394-406">Composition de message</span><span class="sxs-lookup"><span data-stu-id="da394-406">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-407">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-407">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="da394-408">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-408">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="da394-409">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-409">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="da394-410">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-410">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="da394-411">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="da394-411">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="da394-412">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="da394-412">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="da394-413">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="da394-413">Mailbox 1.6</span></span></a></td>
    <td><span data-ttu-id="da394-414">Non disponible</span><span class="sxs-lookup"><span data-stu-id="da394-414">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-415">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="da394-415">Office 2016 for Mac</span></span></td>
    <td> - <span data-ttu-id="da394-416">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="da394-416">Mail Read</span></span><br>
      - <span data-ttu-id="da394-417">Composition de message</span><span class="sxs-lookup"><span data-stu-id="da394-417">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-418">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-418">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="da394-419">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-419">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="da394-420">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-420">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="da394-421">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-421">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="da394-422">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="da394-422">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="da394-423">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="da394-423">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="da394-424">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="da394-424">Mailbox 1.6</span></span></a></td>
    <td><span data-ttu-id="da394-425">Non disponible</span><span class="sxs-lookup"><span data-stu-id="da394-425">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-426">Office 365 pour Android</span><span class="sxs-lookup"><span data-stu-id="da394-426">Office 365 for Android</span></span></td>
    <td> - <span data-ttu-id="da394-427">Lecture de message</span><span class="sxs-lookup"><span data-stu-id="da394-427">Mail Read</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-428">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-428">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="da394-429">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-429">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="da394-430">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-430">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="da394-431">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-431">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="da394-432">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="da394-432">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="da394-433">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="da394-433">Mailbox 1.5</span></span></a></td>
    <td><span data-ttu-id="da394-434">Non disponible</span><span class="sxs-lookup"><span data-stu-id="da394-434">Not available</span></span></td>
  </tr>
</table>

*<span data-ttu-id="da394-435">&ast; : ajouté avec les mises à jour postérieures à la publication.</span><span class="sxs-lookup"><span data-stu-id="da394-435">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="word"></a><span data-ttu-id="da394-436">Word</span><span class="sxs-lookup"><span data-stu-id="da394-436">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="da394-437">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="da394-437">Platform</span></span></th>
    <th><span data-ttu-id="da394-438">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="da394-438">Extension points</span></span></th>
    <th><span data-ttu-id="da394-439">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="da394-439">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="da394-440">API communes</span><span class="sxs-lookup"><span data-stu-id="da394-440">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-441">Office Online</span><span class="sxs-lookup"><span data-stu-id="da394-441">Office Online</span></span></td>
    <td> - <span data-ttu-id="da394-442">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-442">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-443">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-443">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-444">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-444">WordApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-445">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-445">WordApi 1.2</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-446">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-446">WordApi 1.3</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-447">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-447">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="da394-448">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da394-448">BindingEvents</span></span><br>
         - <span data-ttu-id="da394-449">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="da394-449">CustomXmlParts</span></span><br>
         - <span data-ttu-id="da394-450">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-450">DocumentEvents</span></span><br>
         - <span data-ttu-id="da394-451">File</span><span class="sxs-lookup"><span data-stu-id="da394-451">File</span></span><br>
         - <span data-ttu-id="da394-452">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-452">HtmlCoercion</span></span><br>
         - <span data-ttu-id="da394-453">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-453">ImageCoercion</span></span><br>
         - <span data-ttu-id="da394-454">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da394-454">MatrixBindings</span></span><br>
         - <span data-ttu-id="da394-455">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-455">MatrixCoercion</span></span><br>
         - <span data-ttu-id="da394-456">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-456">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="da394-457">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-457">PdfFile</span></span><br>
         - <span data-ttu-id="da394-458">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-458">Selection</span></span><br>
         - <span data-ttu-id="da394-459">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-459">Settings</span></span><br>
         - <span data-ttu-id="da394-460">TableBindings</span><span class="sxs-lookup"><span data-stu-id="da394-460">TableBindings</span></span><br>
         - <span data-ttu-id="da394-461">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-461">TableCoercion</span></span><br>
         - <span data-ttu-id="da394-462">TextBindings</span><span class="sxs-lookup"><span data-stu-id="da394-462">TextBindings</span></span><br>
         - <span data-ttu-id="da394-463">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-463">TextCoercion</span></span><br>
         - <span data-ttu-id="da394-464">TextFile</span><span class="sxs-lookup"><span data-stu-id="da394-464">TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-465">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="da394-465">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="da394-466">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-466">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-467">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-467">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-468">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-468">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-469">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-469">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-470">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-470">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-471">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-471">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="da394-472">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da394-472">BindingEvents</span></span><br>
         - <span data-ttu-id="da394-473">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-473">CompressedFile</span></span><br>
         - <span data-ttu-id="da394-474">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="da394-474">CustomXmlParts</span></span><br>
         - <span data-ttu-id="da394-475">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-475">DocumentEvents</span></span><br>
         - <span data-ttu-id="da394-476">File</span><span class="sxs-lookup"><span data-stu-id="da394-476">File</span></span><br>
         - <span data-ttu-id="da394-477">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-477">HtmlCoercion</span></span><br>
         - <span data-ttu-id="da394-478">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-478">ImageCoercion</span></span><br>
         - <span data-ttu-id="da394-479">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da394-479">MatrixBindings</span></span><br>
         - <span data-ttu-id="da394-480">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-480">MatrixCoercion</span></span><br>
         - <span data-ttu-id="da394-481">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-481">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="da394-482">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-482">PdfFile</span></span><br>
         - <span data-ttu-id="da394-483">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-483">Selection</span></span><br>
         - <span data-ttu-id="da394-484">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-484">Settings</span></span><br>
         - <span data-ttu-id="da394-485">TableBindings</span><span class="sxs-lookup"><span data-stu-id="da394-485">TableBindings</span></span><br>
         - <span data-ttu-id="da394-486">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-486">TableCoercion</span></span><br>
         - <span data-ttu-id="da394-487">TextBindings</span><span class="sxs-lookup"><span data-stu-id="da394-487">TextBindings</span></span><br>
         - <span data-ttu-id="da394-488">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-488">TextCoercion</span></span><br>
         - <span data-ttu-id="da394-489">TextFile</span><span class="sxs-lookup"><span data-stu-id="da394-489">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-490">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="da394-490">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="da394-491">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-491">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-492">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-492">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-493">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-493">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-494">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-494">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-495">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-495">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-496">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-496">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="da394-497">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da394-497">BindingEvents</span></span><br>
         - <span data-ttu-id="da394-498">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-498">CompressedFile</span></span><br>
         - <span data-ttu-id="da394-499">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="da394-499">CustomXmlParts</span></span><br>
         - <span data-ttu-id="da394-500">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-500">DocumentEvents</span></span><br>
         - <span data-ttu-id="da394-501">File</span><span class="sxs-lookup"><span data-stu-id="da394-501">File</span></span><br>
         - <span data-ttu-id="da394-502">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-502">HtmlCoercion</span></span><br>
         - <span data-ttu-id="da394-503">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-503">ImageCoercion</span></span><br>
         - <span data-ttu-id="da394-504">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da394-504">MatrixBindings</span></span><br>
         - <span data-ttu-id="da394-505">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-505">MatrixCoercion</span></span><br>
         - <span data-ttu-id="da394-506">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-506">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="da394-507">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-507">PdfFile</span></span><br>
         - <span data-ttu-id="da394-508">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-508">Selection</span></span><br>
         - <span data-ttu-id="da394-509">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-509">Settings</span></span><br>
         - <span data-ttu-id="da394-510">TableBindings</span><span class="sxs-lookup"><span data-stu-id="da394-510">TableBindings</span></span><br>
         - <span data-ttu-id="da394-511">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-511">TableCoercion</span></span><br>
         - <span data-ttu-id="da394-512">TextBindings</span><span class="sxs-lookup"><span data-stu-id="da394-512">TextBindings</span></span><br>
         - <span data-ttu-id="da394-513">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-513">TextCoercion</span></span><br>
         - <span data-ttu-id="da394-514">TextFile</span><span class="sxs-lookup"><span data-stu-id="da394-514">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-515">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="da394-515">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="da394-516">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-516">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-517">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-517">WordApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-518">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-518">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="da394-519">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da394-519">BindingEvents</span></span><br>
         - <span data-ttu-id="da394-520">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-520">CompressedFile</span></span><br>
         - <span data-ttu-id="da394-521">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="da394-521">CustomXmlParts</span></span><br>
         - <span data-ttu-id="da394-522">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-522">DocumentEvents</span></span><br>
         - <span data-ttu-id="da394-523">File</span><span class="sxs-lookup"><span data-stu-id="da394-523">File</span></span><br>
         - <span data-ttu-id="da394-524">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-524">HtmlCoercion</span></span><br>
         - <span data-ttu-id="da394-525">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-525">ImageCoercion</span></span><br>
         - <span data-ttu-id="da394-526">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da394-526">MatrixBindings</span></span><br>
         - <span data-ttu-id="da394-527">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-527">MatrixCoercion</span></span><br>
         - <span data-ttu-id="da394-528">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-528">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="da394-529">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-529">PdfFile</span></span><br>
         - <span data-ttu-id="da394-530">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-530">Selection</span></span><br>
         - <span data-ttu-id="da394-531">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-531">Settings</span></span><br>
         - <span data-ttu-id="da394-532">TableBindings</span><span class="sxs-lookup"><span data-stu-id="da394-532">TableBindings</span></span><br>
         - <span data-ttu-id="da394-533">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-533">TableCoercion</span></span><br>
         - <span data-ttu-id="da394-534">TextBindings</span><span class="sxs-lookup"><span data-stu-id="da394-534">TextBindings</span></span><br>
         - <span data-ttu-id="da394-535">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-535">TextCoercion</span></span><br>
         - <span data-ttu-id="da394-536">TextFile</span><span class="sxs-lookup"><span data-stu-id="da394-536">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-537">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="da394-537">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="da394-538">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-538">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-539">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-539">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="da394-540">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da394-540">BindingEvents</span></span><br>
         - <span data-ttu-id="da394-541">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-541">CompressedFile</span></span><br>
         - <span data-ttu-id="da394-542">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="da394-542">CustomXmlParts</span></span><br>
         - <span data-ttu-id="da394-543">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-543">DocumentEvents</span></span><br>
         - <span data-ttu-id="da394-544">File</span><span class="sxs-lookup"><span data-stu-id="da394-544">File</span></span><br>
         - <span data-ttu-id="da394-545">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-545">HtmlCoercion</span></span><br>
         - <span data-ttu-id="da394-546">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-546">ImageCoercion</span></span><br>
         - <span data-ttu-id="da394-547">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da394-547">MatrixBindings</span></span><br>
         - <span data-ttu-id="da394-548">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-548">MatrixCoercion</span></span><br>
         - <span data-ttu-id="da394-549">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-549">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="da394-550">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-550">PdfFile</span></span><br>
         - <span data-ttu-id="da394-551">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-551">Selection</span></span><br>
         - <span data-ttu-id="da394-552">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-552">Settings</span></span><br>
         - <span data-ttu-id="da394-553">TableBindings</span><span class="sxs-lookup"><span data-stu-id="da394-553">TableBindings</span></span><br>
         - <span data-ttu-id="da394-554">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-554">TableCoercion</span></span><br>
         - <span data-ttu-id="da394-555">TextBindings</span><span class="sxs-lookup"><span data-stu-id="da394-555">TextBindings</span></span><br>
         - <span data-ttu-id="da394-556">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-556">TextCoercion</span></span><br>
         - <span data-ttu-id="da394-557">TextFile</span><span class="sxs-lookup"><span data-stu-id="da394-557">TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-558">Office 365 pour iPad</span><span class="sxs-lookup"><span data-stu-id="da394-558">Office 365 for iPad</span></span></td>
    <td> - <span data-ttu-id="da394-559">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-559">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-560">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-560">WordApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-561">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-561">WordApi 1.2</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-562">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-562">WordApi 1.3</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-563">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-563">DialogApi 1.1</span></span></a>
</td>
    <td> - <span data-ttu-id="da394-564">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da394-564">BindingEvents</span></span><br>
         - <span data-ttu-id="da394-565">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-565">CompressedFile</span></span><br>
         - <span data-ttu-id="da394-566">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="da394-566">CustomXmlParts</span></span><br>
         - <span data-ttu-id="da394-567">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-567">DocumentEvents</span></span><br>
         - <span data-ttu-id="da394-568">File</span><span class="sxs-lookup"><span data-stu-id="da394-568">File</span></span><br>
         - <span data-ttu-id="da394-569">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-569">HtmlCoercion</span></span><br>
         - <span data-ttu-id="da394-570">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-570">ImageCoercion</span></span><br>
         - <span data-ttu-id="da394-571">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da394-571">MatrixBindings</span></span><br>
         - <span data-ttu-id="da394-572">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-572">MatrixCoercion</span></span><br>
         - <span data-ttu-id="da394-573">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-573">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="da394-574">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-574">PdfFile</span></span><br>
         - <span data-ttu-id="da394-575">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-575">Selection</span></span><br>
         - <span data-ttu-id="da394-576">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-576">Settings</span></span><br>
         - <span data-ttu-id="da394-577">TableBindings</span><span class="sxs-lookup"><span data-stu-id="da394-577">TableBindings</span></span><br>
         - <span data-ttu-id="da394-578">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-578">TableCoercion</span></span><br>
         - <span data-ttu-id="da394-579">TextBindings</span><span class="sxs-lookup"><span data-stu-id="da394-579">TextBindings</span></span><br>
         - <span data-ttu-id="da394-580">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-580">TextCoercion</span></span><br>
         - <span data-ttu-id="da394-581">TextFile</span><span class="sxs-lookup"><span data-stu-id="da394-581">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-582">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="da394-582">Office 365 for Mac</span></span></td>
    <td> - <span data-ttu-id="da394-583">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-583">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-584">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-584">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-585">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-585">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-586">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-586">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-587">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-587">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-588">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-588">DialogApi 1.1</span></span></a>
</td>
    <td> - <span data-ttu-id="da394-589">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da394-589">BindingEvents</span></span><br>
         - <span data-ttu-id="da394-590">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-590">CompressedFile</span></span><br>
         - <span data-ttu-id="da394-591">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="da394-591">CustomXmlParts</span></span><br>
         - <span data-ttu-id="da394-592">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-592">DocumentEvents</span></span><br>
         - <span data-ttu-id="da394-593">File</span><span class="sxs-lookup"><span data-stu-id="da394-593">File</span></span><br>
         - <span data-ttu-id="da394-594">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-594">HtmlCoercion</span></span><br>
         - <span data-ttu-id="da394-595">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-595">ImageCoercion</span></span><br>
         - <span data-ttu-id="da394-596">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da394-596">MatrixBindings</span></span><br>
         - <span data-ttu-id="da394-597">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-597">MatrixCoercion</span></span><br>
         - <span data-ttu-id="da394-598">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-598">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="da394-599">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-599">PdfFile</span></span><br>
         - <span data-ttu-id="da394-600">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-600">Selection</span></span><br>
         - <span data-ttu-id="da394-601">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-601">Settings</span></span><br>
         - <span data-ttu-id="da394-602">TableBindings</span><span class="sxs-lookup"><span data-stu-id="da394-602">TableBindings</span></span><br>
         - <span data-ttu-id="da394-603">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-603">TableCoercion</span></span><br>
         - <span data-ttu-id="da394-604">TextBindings</span><span class="sxs-lookup"><span data-stu-id="da394-604">TextBindings</span></span><br>
         - <span data-ttu-id="da394-605">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-605">TextCoercion</span></span><br>
         - <span data-ttu-id="da394-606">TextFile</span><span class="sxs-lookup"><span data-stu-id="da394-606">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-607">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="da394-607">Office 2019 for Mac</span></span></td>
    <td> - <span data-ttu-id="da394-608">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-608">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-609">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-609">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-610">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-610">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-611">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="da394-611">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-612">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="da394-612">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-613">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-613">DialogApi 1.1</span></span></a>
</td>
    <td> - <span data-ttu-id="da394-614">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da394-614">BindingEvents</span></span><br>
         - <span data-ttu-id="da394-615">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-615">CompressedFile</span></span><br>
         - <span data-ttu-id="da394-616">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="da394-616">CustomXmlParts</span></span><br>
         - <span data-ttu-id="da394-617">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-617">DocumentEvents</span></span><br>
         - <span data-ttu-id="da394-618">File</span><span class="sxs-lookup"><span data-stu-id="da394-618">File</span></span><br>
         - <span data-ttu-id="da394-619">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-619">HtmlCoercion</span></span><br>
         - <span data-ttu-id="da394-620">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-620">ImageCoercion</span></span><br>
         - <span data-ttu-id="da394-621">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da394-621">MatrixBindings</span></span><br>
         - <span data-ttu-id="da394-622">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-622">MatrixCoercion</span></span><br>
         - <span data-ttu-id="da394-623">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-623">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="da394-624">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-624">PdfFile</span></span><br>
         - <span data-ttu-id="da394-625">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-625">Selection</span></span><br>
         - <span data-ttu-id="da394-626">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-626">Settings</span></span><br>
         - <span data-ttu-id="da394-627">TableBindings</span><span class="sxs-lookup"><span data-stu-id="da394-627">TableBindings</span></span><br>
         - <span data-ttu-id="da394-628">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-628">TableCoercion</span></span><br>
         - <span data-ttu-id="da394-629">TextBindings</span><span class="sxs-lookup"><span data-stu-id="da394-629">TextBindings</span></span><br>
         - <span data-ttu-id="da394-630">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-630">TextCoercion</span></span><br>
         - <span data-ttu-id="da394-631">TextFile</span><span class="sxs-lookup"><span data-stu-id="da394-631">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-632">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="da394-632">Office 2016 for Mac</span></span></td>
    <td> - <span data-ttu-id="da394-633">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-633">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="da394-634">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-634">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-635">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-635">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="da394-636">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da394-636">BindingEvents</span></span><br>
         - <span data-ttu-id="da394-637">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-637">CompressedFile</span></span><br>
         - <span data-ttu-id="da394-638">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="da394-638">CustomXmlParts</span></span><br>
         - <span data-ttu-id="da394-639">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-639">DocumentEvents</span></span><br>
         - <span data-ttu-id="da394-640">File</span><span class="sxs-lookup"><span data-stu-id="da394-640">File</span></span><br>
         - <span data-ttu-id="da394-641">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-641">HtmlCoercion</span></span><br>
         - <span data-ttu-id="da394-642">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-642">ImageCoercion</span></span><br>
         - <span data-ttu-id="da394-643">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da394-643">MatrixBindings</span></span><br>
         - <span data-ttu-id="da394-644">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-644">MatrixCoercion</span></span><br>
         - <span data-ttu-id="da394-645">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-645">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="da394-646">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-646">PdfFile</span></span><br>
         - <span data-ttu-id="da394-647">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-647">Selection</span></span><br>
         - <span data-ttu-id="da394-648">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-648">Settings</span></span><br>
         - <span data-ttu-id="da394-649">TableBindings</span><span class="sxs-lookup"><span data-stu-id="da394-649">TableBindings</span></span><br>
         - <span data-ttu-id="da394-650">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-650">TableCoercion</span></span><br>
         - <span data-ttu-id="da394-651">TextBindings</span><span class="sxs-lookup"><span data-stu-id="da394-651">TextBindings</span></span><br>
         - <span data-ttu-id="da394-652">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-652">TextCoercion</span></span><br>
         - <span data-ttu-id="da394-653">TextFile</span><span class="sxs-lookup"><span data-stu-id="da394-653">TextFile</span></span> </td>
  </tr>
</table>

*<span data-ttu-id="da394-654">&ast; : ajouté avec les mises à jour postérieures à la publication.</span><span class="sxs-lookup"><span data-stu-id="da394-654">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="powerpoint"></a><span data-ttu-id="da394-655">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="da394-655">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="da394-656">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="da394-656">Platform</span></span></th>
    <th><span data-ttu-id="da394-657">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="da394-657">Extension points</span></span></th>
    <th><span data-ttu-id="da394-658">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="da394-658">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="da394-659">API communes</span><span class="sxs-lookup"><span data-stu-id="da394-659">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="da394-660">Office Online</span></span></td>
    <td> - <span data-ttu-id="da394-661">Content</span><span class="sxs-lookup"><span data-stu-id="da394-661">Content</span></span><br>
         - <span data-ttu-id="da394-662">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-662">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-663">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-663">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-664">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-664">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="da394-665">ActiveView</span><span class="sxs-lookup"><span data-stu-id="da394-665">ActiveView</span></span><br>
         - <span data-ttu-id="da394-666">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-666">CompressedFile</span></span><br>
         - <span data-ttu-id="da394-667">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-667">DocumentEvents</span></span><br>
         - <span data-ttu-id="da394-668">File</span><span class="sxs-lookup"><span data-stu-id="da394-668">File</span></span><br>
         - <span data-ttu-id="da394-669">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-669">ImageCoercion</span></span><br>
         - <span data-ttu-id="da394-670">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-670">PdfFile</span></span><br>
         - <span data-ttu-id="da394-671">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-671">Selection</span></span><br>
         - <span data-ttu-id="da394-672">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-672">Settings</span></span><br>
         - <span data-ttu-id="da394-673">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-673">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-674">Office 365 pour Windows</span><span class="sxs-lookup"><span data-stu-id="da394-674">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="da394-675">Content</span><span class="sxs-lookup"><span data-stu-id="da394-675">Content</span></span><br>
         - <span data-ttu-id="da394-676">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-676">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-677">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-677">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-678">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-678">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="da394-679">ActiveView</span><span class="sxs-lookup"><span data-stu-id="da394-679">ActiveView</span></span><br>
         - <span data-ttu-id="da394-680">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-680">CompressedFile</span></span><br>
         - <span data-ttu-id="da394-681">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-681">DocumentEvents</span></span><br>
         - <span data-ttu-id="da394-682">File</span><span class="sxs-lookup"><span data-stu-id="da394-682">File</span></span><br>
         - <span data-ttu-id="da394-683">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-683">ImageCoercion</span></span><br>
         - <span data-ttu-id="da394-684">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-684">PdfFile</span></span><br>
         - <span data-ttu-id="da394-685">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-685">Selection</span></span><br>
         - <span data-ttu-id="da394-686">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-686">Settings</span></span><br>
         - <span data-ttu-id="da394-687">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-687">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-688">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="da394-688">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="da394-689">Content</span><span class="sxs-lookup"><span data-stu-id="da394-689">Content</span></span><br>
         - <span data-ttu-id="da394-690">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-690">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-691">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-691">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-692">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-692">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="da394-693">ActiveView</span><span class="sxs-lookup"><span data-stu-id="da394-693">ActiveView</span></span><br>
         - <span data-ttu-id="da394-694">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-694">CompressedFile</span></span><br>
         - <span data-ttu-id="da394-695">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-695">DocumentEvents</span></span><br>
         - <span data-ttu-id="da394-696">File</span><span class="sxs-lookup"><span data-stu-id="da394-696">File</span></span><br>
         - <span data-ttu-id="da394-697">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-697">ImageCoercion</span></span><br>
         - <span data-ttu-id="da394-698">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-698">PdfFile</span></span><br>
         - <span data-ttu-id="da394-699">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-699">Selection</span></span><br>
         - <span data-ttu-id="da394-700">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-700">Settings</span></span><br>
         - <span data-ttu-id="da394-701">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-701">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-702">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="da394-702">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="da394-703">Content</span><span class="sxs-lookup"><span data-stu-id="da394-703">Content</span></span><br>
         - <span data-ttu-id="da394-704">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-704">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-705">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-705">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="da394-706">ActiveView</span><span class="sxs-lookup"><span data-stu-id="da394-706">ActiveView</span></span><br>
         - <span data-ttu-id="da394-707">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-707">CompressedFile</span></span><br>
         - <span data-ttu-id="da394-708">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-708">DocumentEvents</span></span><br>
         - <span data-ttu-id="da394-709">File</span><span class="sxs-lookup"><span data-stu-id="da394-709">File</span></span><br>
         - <span data-ttu-id="da394-710">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-710">ImageCoercion</span></span><br>
         - <span data-ttu-id="da394-711">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-711">PdfFile</span></span><br>
         - <span data-ttu-id="da394-712">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-712">Selection</span></span><br>
         - <span data-ttu-id="da394-713">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-713">Settings</span></span><br>
         - <span data-ttu-id="da394-714">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-714">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-715">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="da394-715">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="da394-716">Content</span><span class="sxs-lookup"><span data-stu-id="da394-716">Content</span></span><br>
         - <span data-ttu-id="da394-717">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-717">TaskPane</span></span><br>
    </td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-718">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-718">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="da394-719">ActiveView</span><span class="sxs-lookup"><span data-stu-id="da394-719">ActiveView</span></span><br>
         - <span data-ttu-id="da394-720">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-720">CompressedFile</span></span><br>
         - <span data-ttu-id="da394-721">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-721">DocumentEvents</span></span><br>
         - <span data-ttu-id="da394-722">File</span><span class="sxs-lookup"><span data-stu-id="da394-722">File</span></span><br>
         - <span data-ttu-id="da394-723">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-723">ImageCoercion</span></span><br>
         - <span data-ttu-id="da394-724">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-724">PdfFile</span></span><br>
         - <span data-ttu-id="da394-725">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-725">Selection</span></span><br>
         - <span data-ttu-id="da394-726">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-726">Settings</span></span><br>
         - <span data-ttu-id="da394-727">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-727">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-728">Office 365 pour iPad</span><span class="sxs-lookup"><span data-stu-id="da394-728">Office 365 for iPad</span></span></td>
    <td> - <span data-ttu-id="da394-729">Content</span><span class="sxs-lookup"><span data-stu-id="da394-729">Content</span></span><br>
         - <span data-ttu-id="da394-730">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-730">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-731">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-731">DialogApi 1.1</span></span></a></td>
     <td> - <span data-ttu-id="da394-732">ActiveView</span><span class="sxs-lookup"><span data-stu-id="da394-732">ActiveView</span></span><br>
         - <span data-ttu-id="da394-733">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-733">CompressedFile</span></span><br>
         - <span data-ttu-id="da394-734">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-734">DocumentEvents</span></span><br>
         - <span data-ttu-id="da394-735">File</span><span class="sxs-lookup"><span data-stu-id="da394-735">File</span></span><br>
         - <span data-ttu-id="da394-736">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-736">PdfFile</span></span><br>
         - <span data-ttu-id="da394-737">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-737">Selection</span></span><br>
         - <span data-ttu-id="da394-738">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-738">Settings</span></span><br>
         - <span data-ttu-id="da394-739">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-739">TextCoercion</span></span><br>
         - <span data-ttu-id="da394-740">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-740">ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-741">Office 365 pour Mac</span><span class="sxs-lookup"><span data-stu-id="da394-741">Office 365 for Mac</span></span></td>
    <td> - <span data-ttu-id="da394-742">Content</span><span class="sxs-lookup"><span data-stu-id="da394-742">Content</span></span><br>
         - <span data-ttu-id="da394-743">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-743">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-744">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-744">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-745">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-745">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="da394-746">ActiveView</span><span class="sxs-lookup"><span data-stu-id="da394-746">ActiveView</span></span><br>
         - <span data-ttu-id="da394-747">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-747">CompressedFile</span></span><br>
         - <span data-ttu-id="da394-748">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-748">DocumentEvents</span></span><br>
         - <span data-ttu-id="da394-749">File</span><span class="sxs-lookup"><span data-stu-id="da394-749">File</span></span><br>
         - <span data-ttu-id="da394-750">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-750">ImageCoercion</span></span><br>
         - <span data-ttu-id="da394-751">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-751">PdfFile</span></span><br>
         - <span data-ttu-id="da394-752">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-752">Selection</span></span><br>
         - <span data-ttu-id="da394-753">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-753">Settings</span></span><br>
         - <span data-ttu-id="da394-754">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-754">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-755">Office 2019 pour Mac</span><span class="sxs-lookup"><span data-stu-id="da394-755">Office 2019 for Mac</span></span></td>
    <td> - <span data-ttu-id="da394-756">Content</span><span class="sxs-lookup"><span data-stu-id="da394-756">Content</span></span><br>
         - <span data-ttu-id="da394-757">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-757">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-758">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-758">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-759">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-759">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="da394-760">ActiveView</span><span class="sxs-lookup"><span data-stu-id="da394-760">ActiveView</span></span><br>
         - <span data-ttu-id="da394-761">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-761">CompressedFile</span></span><br>
         - <span data-ttu-id="da394-762">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-762">DocumentEvents</span></span><br>
         - <span data-ttu-id="da394-763">File</span><span class="sxs-lookup"><span data-stu-id="da394-763">File</span></span><br>
         - <span data-ttu-id="da394-764">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-764">ImageCoercion</span></span><br>
         - <span data-ttu-id="da394-765">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-765">PdfFile</span></span><br>
         - <span data-ttu-id="da394-766">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-766">Selection</span></span><br>
         - <span data-ttu-id="da394-767">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-767">Settings</span></span><br>
         - <span data-ttu-id="da394-768">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-768">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-769">Office 2016 pour Mac</span><span class="sxs-lookup"><span data-stu-id="da394-769">Office 2016 for Mac</span></span></td>
    <td> - <span data-ttu-id="da394-770">Content</span><span class="sxs-lookup"><span data-stu-id="da394-770">Content</span></span><br>
         - <span data-ttu-id="da394-771">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-771">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-772">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-772">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="da394-773">ActiveView</span><span class="sxs-lookup"><span data-stu-id="da394-773">ActiveView</span></span><br>
         - <span data-ttu-id="da394-774">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da394-774">CompressedFile</span></span><br>
         - <span data-ttu-id="da394-775">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-775">DocumentEvents</span></span><br>
         - <span data-ttu-id="da394-776">File</span><span class="sxs-lookup"><span data-stu-id="da394-776">File</span></span><br>
         - <span data-ttu-id="da394-777">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-777">ImageCoercion</span></span><br>
         - <span data-ttu-id="da394-778">PdfFile</span><span class="sxs-lookup"><span data-stu-id="da394-778">PdfFile</span></span><br>
         - <span data-ttu-id="da394-779">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-779">Selection</span></span><br>
         - <span data-ttu-id="da394-780">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-780">Settings</span></span><br>
         - <span data-ttu-id="da394-781">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-781">TextCoercion</span></span></td>
  </tr>
</table>

*<span data-ttu-id="da394-782">&ast; : ajouté avec les mises à jour postérieures à la publication.</span><span class="sxs-lookup"><span data-stu-id="da394-782">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="onenote"></a><span data-ttu-id="da394-783">OneNote</span><span class="sxs-lookup"><span data-stu-id="da394-783">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="da394-784">Plate-forme</span><span class="sxs-lookup"><span data-stu-id="da394-784">Platform</span></span></th>
    <th><span data-ttu-id="da394-785">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="da394-785">Extension points</span></span></th>
    <th><span data-ttu-id="da394-786">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="da394-786">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="da394-787">API communes</span><span class="sxs-lookup"><span data-stu-id="da394-787">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-788">Office Online</span><span class="sxs-lookup"><span data-stu-id="da394-788">Office Online</span></span></td>
    <td> - <span data-ttu-id="da394-789">Content</span><span class="sxs-lookup"><span data-stu-id="da394-789">Content</span></span><br>
         - <span data-ttu-id="da394-790">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-790">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="da394-791">Commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-791">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets"><span data-ttu-id="da394-792">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-792">OneNoteApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-793">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-793">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="da394-794">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da394-794">DocumentEvents</span></span><br>
         - <span data-ttu-id="da394-795">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-795">HtmlCoercion</span></span><br>
         - <span data-ttu-id="da394-796">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-796">ImageCoercion</span></span><br>
         - <span data-ttu-id="da394-797">Settings</span><span class="sxs-lookup"><span data-stu-id="da394-797">Settings</span></span><br>
         - <span data-ttu-id="da394-798">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-798">TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="da394-799">Project</span><span class="sxs-lookup"><span data-stu-id="da394-799">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="da394-800">Plateforme</span><span class="sxs-lookup"><span data-stu-id="da394-800">Platform</span></span></th>
    <th><span data-ttu-id="da394-801">Points d’extension</span><span class="sxs-lookup"><span data-stu-id="da394-801">Extension points</span></span></th>
    <th><span data-ttu-id="da394-802">Ensembles de conditions requises de l’API</span><span class="sxs-lookup"><span data-stu-id="da394-802">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="da394-803">API communes</span><span class="sxs-lookup"><span data-stu-id="da394-803">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-804">Office 2019 pour Windows</span><span class="sxs-lookup"><span data-stu-id="da394-804">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="da394-805">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-805">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-806">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-806">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="da394-807">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-807">Selection</span></span><br>
         - <span data-ttu-id="da394-808">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-808">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-809">Office 2016 pour Windows</span><span class="sxs-lookup"><span data-stu-id="da394-809">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="da394-810">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-810">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-811">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-811">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="da394-812">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-812">Selection</span></span><br>
         - <span data-ttu-id="da394-813">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-813">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da394-814">Office 2013 pour Windows</span><span class="sxs-lookup"><span data-stu-id="da394-814">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="da394-815">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da394-815">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="da394-816">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="da394-816">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="da394-817">Selection</span><span class="sxs-lookup"><span data-stu-id="da394-817">Selection</span></span><br>
         - <span data-ttu-id="da394-818">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da394-818">TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="da394-819">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="da394-819">See also</span></span>

- [<span data-ttu-id="da394-820">Vue d’ensemble de la plateforme de compléments pour Office</span><span class="sxs-lookup"><span data-stu-id="da394-820">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="da394-821">Ensembles de conditions requises des API communes</span><span class="sxs-lookup"><span data-stu-id="da394-821">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="da394-822">Ensembles de conditions requises concernant des commandes de complément</span><span class="sxs-lookup"><span data-stu-id="da394-822">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="da394-823">Documentation de référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="da394-823">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="da394-824">Historique des mises à jour d’Office 2016 et 2019 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="da394-824">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="da394-825">Historique des mises à jour d’Office 2013 (Démarrer en un clic)</span><span class="sxs-lookup"><span data-stu-id="da394-825">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="da394-826">Historique des mises à jour d’Office 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="da394-826">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="da394-827">Historique des mises à jour d’Outlook 2010, 2013 et 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="da394-827">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)